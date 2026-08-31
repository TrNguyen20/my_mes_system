from flask import Flask, render_template, request, redirect, Response, flash, url_for
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from flask_bcrypt import Bcrypt
from datetime import datetime
from functools import wraps
import csv
import io

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///mes_v2.db'
app.config['SECRET_KEY'] = 'mes_secret_key_123'

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
    date_created = db.Column(db.DateTime, default=datetime.now)

    def set_password(self, password):
        self.password_hash = bcrypt.generate_password_hash(password).decode('utf-8')

    def check_password(self, password):
        return bcrypt.check_password_hash(self.password_hash, password)

class WorkOrder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    po_number = db.Column(db.String(50), nullable=False)
    part_name = db.Column(db.String(100), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), default='Pending')
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

# --- HELPERS ---

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if not current_user.is_authenticated or current_user.role not in roles:
                flash("Bạn không có quyền thực hiện hành động này!", "danger")
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

def log_action(action, target_type, target_id, detail=''):
    log = AuditLog(user_id=current_user.id, action=action, target_type=target_type, target_id=target_id, detail=detail)
    db.session.add(log)
    db.session.commit()

# --- ROUTES ---

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = User.query.filter_by(username=request.form['username']).first()
        if user and user.check_password(request.form['password']):
            login_user(user, remember=request.form.get('remember'))
            return redirect(url_for('index'))
        flash('Sai tài khoản hoặc mật khẩu!', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    search = request.args.get('search')
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
    new_order = WorkOrder(
        po_number=request.form['po_number'],
        part_name=request.form['part_name'],
        quantity=request.form['quantity']
    )
    db.session.add(new_order)
    db.session.commit()
    log_action('CREATE', 'WorkOrder', new_order.id, f'PO: {new_order.po_number}')
    flash('Đã tạo lệnh sản xuất mới!', 'success')
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

@app.route('/delete/<int:id>')
@login_required
@role_required('admin')
def delete_order(id):
    order = WorkOrder.query.get_or_404(id)
    po_ref = order.po_number
    db.session.delete(order)
    db.session.commit()
    log_action('DELETE', 'WorkOrder', id, f'PO: {po_ref}')
    flash(f'Đã xóa lệnh {po_ref}', 'warning')
    return redirect(url_for('index'))

@app.route('/export')
@login_required
@role_required('admin', 'manager')
def export_csv():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['PO Number', 'Part Name', 'Qty', 'Status', 'Duration (Min)'])
    orders = WorkOrder.query.all()
    for o in orders:
        writer.writerow([o.po_number, o.part_name, o.quantity, o.status, o.get_duration()])
    output.seek(0)
    log_action('EXPORT', 'System', 0, 'Exported WorkOrders CSV')
    return Response(output, mimetype="text/csv", headers={"Content-Disposition": "attachment;filename=mes_report.csv"})

@app.route('/audit-log')
@login_required
@role_required('admin')
def view_audit_log():
    logs = AuditLog.query.order_by(AuditLog.timestamp.desc()).limit(100).all()
    return render_template('audit_log.html', logs=logs)

# --- CLI COMMANDS ---

@app.cli.command("seed-admin")
def seed_admin():
    db.create_all()
    if not User.query.filter_by(username='admin').first():
        admin = User(username='admin', email='admin@factory.com', role='admin')
        admin.set_password('admin123')
        db.session.add(admin)
        db.session.commit()
        print("Admin created: admin / admin123")

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
    app.run(debug=True)