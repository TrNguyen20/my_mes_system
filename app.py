import io
import csv
import openpyxl
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, Response, flash, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from flask_bcrypt import Bcrypt
from openpyxl.styles import Font, PatternFill, Alignment


app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///mes_v2.db'
app.config['SECRET_KEY'] = 'industrial_level_secret'

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# --- MODELS ---

class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(128), nullable=False)
    role = db.Column(db.String(20), default='operator') # admin, manager, operator
    is_active = db.Column(db.Boolean, default=True)
    date_created = db.Column(db.DateTime, default=datetime.now)

    def set_password(self, password):
        self.password_hash = bcrypt.generate_password_hash(password).decode('utf-8')

    def check_password(self, password):
        return bcrypt.check_password_hash(self.password_hash, password)

class WorkOrder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    po_number = db.Column(db.String(50), unique=True, nullable=False)
    part_name = db.Column(db.String(100), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), default='Pending') # Pending, In Progress, Completed
    date_created = db.Column(db.DateTime, default=datetime.now)
    date_started = db.Column(db.DateTime, nullable=True)
    date_completed = db.Column(db.DateTime, nullable=True)

    def get_duration(self):
        if self.date_started and self.date_completed:
            diff = self.date_completed - self.date_started
            return round(diff.total_seconds() / 60, 2)
        return None

class AuditLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    action = db.Column(db.String(50), nullable=False)
    target_type = db.Column(db.String(50), nullable=False)
    target_id = db.Column(db.Integer, nullable=False)
    detail = db.Column(db.String(200))
    timestamp = db.Column(db.DateTime, default=datetime.now)
    user = db.relationship('User', backref='logs')

# --- HELPERS & DECORATORS ---

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not current_user.is_authenticated or current_user.role not in roles:
                flash("Quyền truy cập bị từ chối!", "danger")
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

def log_action(action, target_type, target_id, detail=''):
    log = AuditLog(user_id=current_user.id, action=action, target_type=target_type, target_id=target_id, detail=detail)
    db.session.add(log)
    db.session.commit()

# --- AUTH ROUTES ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = User.query.filter_by(username=request.form['username']).first()
        if user and user.check_password(request.form['password']):
            if not user.is_active:
                flash("Tài khoản đã bị vô hiệu hóa!", "danger")
                return redirect(url_for('login'))
            login_user(user, remember=request.form.get('remember'))
            return redirect(url_for('index'))
        flash('Sai tên đăng nhập hoặc mật khẩu!', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

# --- CORE MES ROUTES ---

@app.route('/')
@login_required
def index():
    search = request.args.get('search', '')
    query = WorkOrder.query
    if search:
        query = query.filter(WorkOrder.po_number.contains(search) | WorkOrder.part_name.contains(search))
    orders = query.order_by(WorkOrder.date_created.desc()).all()

    stats = {
        'pending': WorkOrder.query.filter_by(status='Pending').count(),
        'progress': WorkOrder.query.filter_by(status='In Progress').count(),
        'completed': WorkOrder.query.filter_by(status='Completed').count(),
    }
    return render_template('index.html', orders=orders, stats=stats)

@app.route('/add', methods=['POST'])
@login_required
@role_required('admin', 'manager')
def add_order():
    po = request.form['po_number'].strip()
    if WorkOrder.query.filter_by(po_number=po).first():
        flash(f"Lỗi: Mã PO {po} đã tồn tại!", "danger")
        return redirect(url_for('index'))
    
    new_order = WorkOrder(
        po_number=po,
        part_name=request.form['part_name'],
        quantity=request.form['quantity']
    )
    db.session.add(new_order)
    db.session.commit()
    log_action('CREATE', 'WorkOrder', new_order.id, f'PO: {new_order.po_number}')
    flash("Tạo lệnh sản xuất thành công!", "success")
    return redirect(url_for('index'))

@app.route('/update/<int:id>')
@login_required
def update_status(id):
    order = WorkOrder.query.get_or_404(id)
    old_status = order.status
    if order.status == 'Pending':
        order.status = 'In Progress'
        order.date_started = datetime.now()
    elif order.status == 'In Progress':
        order.status = 'Completed'
        order.date_completed = datetime.now()
    db.session.commit()
    log_action('STATUS_CHANGE', 'WorkOrder', order.id, f'{old_status} -> {order.status}')
    return redirect(url_for('index'))

# --- EXCEL & EXPORT ROUTES ---

@app.route('/import-excel', methods=['POST'])
@login_required
@role_required('admin', 'manager')
def import_excel():
    file = request.files.get('excel_file')
    if not file or not file.filename.endswith('.xlsx'):
        flash("File không hợp lệ!", "danger")
        return redirect(url_for('index'))

    wb = openpyxl.load_workbook(file)
    ws = wb.active
    imported, skipped, invalid = 0, 0, 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        po, name, qty = str(row[0]).strip(), str(row[1]).strip(), row[2]
        if not po or po == "None": continue
        
        if WorkOrder.query.filter_by(po_number=po).first():
            skipped += 1
            continue
        try:
            qty_i = int(qty)
            if qty_i <= 0: raise ValueError
        except:
            invalid += 1
            continue

        db.session.add(WorkOrder(po_number=po, part_name=name, quantity=qty_i))
        imported += 1
    
    db.session.commit()
    log_action('IMPORT_EXCEL', 'System', 0, f'Success: {imported}')
    flash(f"✅ Đã nhập: {imported} | ⏭ Trùng: {skipped} | ❌ Lỗi: {invalid}", "info")
    return redirect(url_for('index'))

@app.route('/download-excel-template')
def download_template():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Import Template"
    
    # Styling
    header_fill = PatternFill(start_color="2d3436", end_color="2d3436", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    headers = ["PO Number", "Part Name", "Quantity"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 15
    
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="MES_Template.xlsx")

#Để Excel nhận diện đúng tiếng Việt hoặc ký tự đặc biệt, ta cần thêm \ufeff vào đầu file.
@app.route('/export-csv')
@login_required
@role_required('admin', 'manager')
def export_csv():
    output = io.StringIO()
    # Thêm ký tự BOM để Excel mở đúng định dạng UTF-8
    output.write('\ufeff') 
    writer = csv.writer(output)
    writer.writerow(['PO Number', 'Tên sản phẩm', 'Số lượng', 'Trạng thái', 'Lead Time (Phút)'])
    
    orders = WorkOrder.query.all()
    for o in orders:
        writer.writerow([o.po_number, o.part_name, o.quantity, o.status, o.get_duration()])
    
    output.seek(0)
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment;filename=report_mes.csv"}
    )

# --- ADMIN ROUTES ---

@app.route('/admin/users')
@login_required
@role_required('admin')
def admin_users():
    users = User.query.all()
    return render_template('admin_users.html', users=users)

@app.route('/admin/create-user', methods=['POST'])
@login_required
@role_required('admin')
def admin_create_user():
    user = User(username=request.form['username'], email=request.form['email'], role=request.form['role'])
    user.set_password(request.form['password'])
    db.session.add(user)
    db.session.commit()
    flash("Tạo người dùng thành công!", "success")
    return redirect(url_for('admin_users'))

@app.route('/admin/toggle-user/<int:id>', methods=['POST'])
@login_required
@role_required('admin')
def admin_toggle_user(id):
    if current_user.id == id: return redirect(url_for('admin_users'))
    u = User.query.get_or_404(id)
    u.is_active = not u.is_active
    db.session.commit()
    return redirect(url_for('admin_users'))

@app.route('/audit-log')
@login_required
@role_required('admin')
def view_audit_log():
    logs = AuditLog.query.order_by(AuditLog.timestamp.desc()).limit(100).all()
    return render_template('audit_log.html', logs=logs)

# --- INIT ---

@app.cli.command("seed-admin")
def seed_admin():
    db.create_all()
    if not User.query.filter_by(username='admin').first():
        u = User(username='admin', email='admin@mes.com', role='admin')
        u.set_password('admin123')
        db.session.add(u)
        db.session.commit()
        print("Admin account created: admin / admin123")

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True)


from openpyxl.styles import Font, PatternFill, Alignment

@app.route('/download-excel-template')
def download_template():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Import Template"

    # Định nghĩa các Styles
    header_style = PatternFill(start_color="2d3436", end_color="2d3436", fill_type="solid")
    sample_style = PatternFill(start_color="dfe6e9", end_color="dfe6e9", fill_type="solid")
    white_font = Font(bold=True, color="FFFFFF")
    sample_font = Font(color="2d3436")
    warning_font = Font(italic=True, color="e17055") # Màu cam (orange)
    center_align = Alignment(horizontal="center", vertical="center")

    # 1. Header Row (Dòng 1)
    headers = ["PO Number", "Part Name", "Quantity"]
    ws.append(headers)
    ws.row_dimensions[1].height = 20
    for cell in ws[1]:
        cell.fill = header_style
        cell.font = white_font
        cell.alignment = center_align

    # 2. Note Row (Dòng 6)
    ws.merge_cells('A6:C6')
    ws['A6'] = "⚠ Delete sample rows before importing. Keep header row."
    ws['A6'].font = warning_font
    ws['A6'].alignment = center_align

    # 3. Sample Data Rows (Dòng 20 - 25)
    samples = [
        ["PO-SAMPLE-001", "Engine Bolt M8", 500],
        ["PO-SAMPLE-002", "Shaft Bearing 6205", 200],
        ["PO-SAMPLE-003", "Cover Plate A3", 150],
        ["PO-SAMPLE-004", "Example Item 4", 100],
        ["PO-SAMPLE-005", "Example Item 5", 300],
        ["PO-SAMPLE-025", "Cavity Plate A3", 150]
    ]]
    
    # Ghi dữ liệu mẫu vào đúng dòng yêu cầu
    current_row = 20
    for data in samples:
        for col, value in enumerate(data, 1):
            cell = ws.cell(row=current_row, column=col, value=value)
            cell.fill = sample_style
            cell.font = sample_font
            cell.alignment = center_align
        current_row += 1

    # 4. Column Widths
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 15

    # 5. Freeze Pane (Cố định dòng 1)
    ws.freeze_panes = "A2"

    wb.save(output)
    output.seek(0)
    return send_file(
        output, 
        as_attachment=True, 
        download_name="MES_Import_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )