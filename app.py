from flask import Flask, render_template, request, redirect, Response
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import csv
import io

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///mes_v2.db'
db = SQLAlchemy(app)

class WorkOrder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    po_number = db.Column(db.String(50), nullable=False) # New: PO Tracking
    part_name = db.Column(db.String(100), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), default='Pending')
    date_created = db.Column(db.DateTime, default=datetime.now)
    date_started = db.Column(db.DateTime, nullable=True)
    date_completed = db.Column(db.DateTime, nullable=True)

    # Logic to calculate duration in minutes
    def get_duration(self):
        if self.date_started and self.date_completed:
            diff = self.date_completed - self.date_started
            return round(diff.total_seconds() / 60, 2) # Returns minutes
        return None

with app.app_context():
    db.create_all()

@app.route('/')
def index():
    # Search functionality for Managers
    search = request.args.get('search')
    if search:
        orders = WorkOrder.query.filter(WorkOrder.po_number.contains(search) | WorkOrder.part_name.contains(search)).all()
    else:
        orders = WorkOrder.query.order_by(WorkOrder.date_created.desc()).all()

    # Dashboard Metrics
    stats = {
        'pending': WorkOrder.query.filter_by(status='Pending').count(),
        'progress': WorkOrder.query.filter_by(status='In Progress').count(),
        'completed': WorkOrder.query.filter_by(status='Completed').count(),
    }
    return render_template('index.html', orders=orders, stats=stats)

@app.route('/add', methods=['POST'])
def add_order():
    new_order = WorkOrder(
        po_number=request.form['po_number'],
        part_name=request.form['part_name'],
        quantity=request.form['quantity']
    )
    db.session.add(new_order)
    db.session.commit()
    return redirect('/')

@app.route('/update/<int:id>')
def update_status(id):
    order = WorkOrder.query.get_or_404(id)
    if order.status == 'Pending':
        order.status = 'In Progress'
        order.date_started = datetime.now()
    elif order.status == 'In Progress':
        order.status = 'Completed'
        order.date_completed = datetime.now()
    db.session.commit()
    return redirect('/')

# CSV EXPORT FEATURE
@app.route('/export')
def export_csv():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['PO Number', 'Part Name', 'Qty', 'Status', 'Duration (Min)'])
    
    orders = WorkOrder.query.all()
    for o in orders:
        writer.writerow([o.po_number, o.part_name, o.quantity, o.status, o.get_duration()])
    
    output.seek(0)
    return Response(output, mimetype="text/csv", headers={"Content-Disposition": "attachment;filename=factory_report.csv"})

if __name__ == "__main__":
    app.run(debug=True)