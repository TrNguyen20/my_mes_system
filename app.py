import io
import csv
import openpyxl
import gspread
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, request, redirect, Response, flash, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from flask_bcrypt import Bcrypt
from openpyxl.styles import Font, PatternFill, Alignment
from sqlalchemy import or_

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///mes_v3.db'
app.config['SECRET_KEY'] = 'enterprise_mes_routing_2025'

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# --- DATABASE MODELS ---

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
    status = db.Column(db.String(20), default='Pending')
    current_step_index = db.Column(db.Integer, default=0)
    date_created = db.Column(db.DateTime, default=datetime.now)
    date_started = db.Column(db.DateTime, nullable=True)
    date_completed = db.Column(db.DateTime, nullable=True)
    operations = db.relationship('Operation', backref='work_order', lazy=True, cascade="all, delete-orphan")

    def get_duration(self):
        if self.date_started and self.date_completed:
            diff = self.date_completed - self.date_started
            return round(diff.total_seconds() / 60, 2)
        return None

class Operation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    work_order_id = db.Column(db.Integer, db.ForeignKey('work_order.id'), nullable=False)
    name = db.Column(db.String(100), nullable=False)
    sequence = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), default='Pending')

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

# --- ROUTES ---

@app.route('/')
@login_required
def index():
    search = request.args.get('search', '')
    query = WorkOrder.query
    if search:
        query = query.filter(or_(
            WorkOrder.po_number.contains(search),
            WorkOrder.part_name.contains(search),
            WorkOrder.status.contains(search),
            db.cast(WorkOrder.quantity, db.String).contains(search)
        ))
    orders = query.order_by(WorkOrder.date_created.desc()).all()
    stats = {
        'pending': WorkOrder.query.filter_by(status='Pending').count(),
        'progress': WorkOrder.query.filter_by(status='In Progress').count(),
        'completed': WorkOrder.query.filter_by(status='Completed').count(),
    }
    return render_template('index.html', orders=orders, stats=stats, search_val=search)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = User.query.filter_by(username=request.form['username']).first()
        if user and user.check_password(request.form['password']):
            if not user.is_active:
                flash("Tài khoản đã bị khóa!", "danger")
                return redirect(url_for('login'))
            login_user(user, remember=request.form.get('remember'))
            return redirect(url_for('index'))
        flash('Sai tài khoản hoặc mật khẩu!', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

# --- PRODUCTION LOGIC ---

@app.route('/add', methods=['POST'])
@login_required
@role_required('admin', 'manager')
def add_order():
    po = request.form['po_number'].strip()
    if WorkOrder.query.filter_by(po_number=po).first():
        flash(f"Lỗi: PO {po} đã tồn tại!", "danger")
        return redirect(url_for('index'))
    
    new_order = WorkOrder(po_number=po, part_name=request.form['part_name'], quantity=request.form['quantity'])
    db.session.add(new_order)
    db.session.flush()

    ops_list = request.form.getlist('operations[]')
    if not ops_list or ops_list[0] == "": ops_list = ["Sản xuất mặc định"]

    for i, op_name in enumerate(ops_list):
        db.session.add(Operation(work_order_id=new_order.id, name=op_name, sequence=i+1))
    
    db.session.commit()
    log_action('CREATE_ROUTING', 'WorkOrder', new_order.id, f'PO: {po}')
    flash("Đã mở lệnh sản xuất mới!", "success")
    return redirect(url_for('index'))

@app.route('/edit/<int:id>', methods=['POST'])
@login_required
@role_required('admin', 'manager')
def edit_order(id):
    order = WorkOrder.query.get_or_404(id)
    order.po_number = request.form['po_number'].strip()
    order.part_name = request.form['part_name'].strip()
    order.quantity = request.form['quantity']
    db.session.commit()
    log_action('EDIT', 'WorkOrder', id, f'Edited PO: {order.po_number}')
    flash("Đã cập nhật thông tin đơn hàng!", "success")
    return redirect(url_for('index'))

@app.route('/update/<int:id>')
@login_required
def update_status(id):
    order = WorkOrder.query.get_or_404(id)
    ops = sorted(order.operations, key=lambda x: x.sequence)
    
    if order.status == 'Pending':
        order.status = 'In Progress'
        order.date_started = datetime.now()
        order.current_step_index = 1
    elif order.status == 'In Progress':
        ops[order.current_step_index - 1].status = 'Completed'
        if order.current_step_index < len(ops):
            order.current_step_index += 1
        else:
            order.status = 'Completed'
            order.date_completed = datetime.now()
    db.session.commit()
    return redirect(url_for('index'))

# --- SYNC & DATA ---

@app.route('/import-from-sheets')
@login_required
@role_required('admin', 'manager')
def import_from_sheets():
    # ID Sheet từ đường dẫn bạn cung cấp
    SHEET_ID = "1VeoJY4tW3EoN-IK_kXKGlExH_E1Jzr6ffQB-eVbJO1w"
    try:
        gc = gspread.service_account(filename='credentials.json')
        sh = gc.open_by_key(SHEET_ID)
        worksheet = sh.get_worksheet(0)
        records = worksheet.get_all_records()
        
        imported = 0
        for row in records:
            po = str(row.get('po_number')).strip()
            if po and not WorkOrder.query.filter_by(po_number=po).first():
                new_order = WorkOrder(po_number=po, part_name=row.get('part_name'), quantity=int(row.get('quantity', 0)))
                db.session.add(new_order)
                db.session.flush()
                db.session.add(Operation(work_order_id=new_order.id, name="Sản xuất (Sheets)", sequence=1))
                imported += 1
        db.session.commit()
        flash(f"📊 Đồng bộ Google Sheets thành công: +{imported} đơn mới", "success")
    except Exception as e:
        flash(f"Lỗi đồng bộ: {str(e)}", "danger")
    return redirect(url_for('index'))

@app.route('/export-csv')
@login_required
@role_required('admin', 'manager')
def export_csv():
    output = io.StringIO()
    output.write('\ufeff')
    writer = csv.writer(output)
    writer.writerow(['PO Number', 'Tên sản phẩm', 'Số lượng', 'Trạng thái', 'Lead Time'])
    for o in WorkOrder.query.all():
        writer.writerow([o.po_number, o.part_name, o.quantity, o.status, o.get_duration()])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition":"attachment;filename=report.csv"})

# --- ADMIN & LOGS ---

@app.route('/admin/users')
@login_required
@role_required('admin')
def admin_users():
    users = User.query.all()
    user_count = User.query.count()
    return render_template('admin_users.html', users=users, user_count=user_count)

@app.route('/admin/toggle-user/<int:id>', methods=['POST'])
@login_required
@role_required('admin')
def admin_toggle_user(id):
    if current_user.id != id:
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

@app.cli.command("seed-admin")
def seed_admin():
    db.create_all()
    if not User.query.filter_by(username='admin').first():
        u = User(username='admin', email='admin@mes.com', role='admin')
        u.set_password('admin123')
        db.session.add(u)
        db.session.commit()
        print("Admin account: admin / admin123")

if __name__ == "__main__":
    with app.app_context(): db.create_all()
    app.run(debug=True)