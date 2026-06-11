from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
from functools import wraps
import os
from io import BytesIO
from openpyxl import Workbook
import secrets

app = Flask(__name__)
app.config['SECRET_KEY'] = secrets.token_hex(16)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///requisition_system.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# ==================== DATABASE MODELS ====================

class Department(db.Model):
    __tablename__ = 'departments'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), unique=True, nullable=False)
    description = db.Column(db.String(200))
    
    users = db.relationship('User', backref='department_ref', foreign_keys='User.department_id')
    supervisors = db.relationship('DepartmentSupervisor', backref='department', cascade='all, delete-orphan')

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(50), nullable=False)
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'), nullable=True)
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    requisitions = db.relationship('Requisition', foreign_keys='Requisition.requestor_id', backref='requestor')
    approvals_given = db.relationship('Approval', foreign_keys='Approval.approver_id', backref='approver')

class DepartmentSupervisor(db.Model):
    __tablename__ = 'department_supervisors'
    id = db.Column(db.Integer, primary_key=True)
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'), nullable=False)
    supervisor_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    
    supervisor = db.relationship('User', foreign_keys=[supervisor_id])

class Requisition(db.Model):
    __tablename__ = 'requisitions'
    id = db.Column(db.Integer, primary_key=True)
    requestor_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    department_id = db.Column(db.Integer, db.ForeignKey('departments.id'), nullable=False)
    reason = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(50), default='Pending Supervisor Approval')
    current_approval_level = db.Column(db.Integer, default=1)
    date_created = db.Column(db.DateTime, default=datetime.utcnow)
    date_updated = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    items = db.relationship('RequisitionItem', backref='requisition', cascade='all, delete-orphan')
    approvals = db.relationship('Approval', backref='requisition', cascade='all, delete-orphan')

class RequisitionItem(db.Model):
    __tablename__ = 'requisition_items'
    id = db.Column(db.Integer, primary_key=True)
    requisition_id = db.Column(db.Integer, db.ForeignKey('requisitions.id'), nullable=False)
    description = db.Column(db.String(200), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    estimated_cost = db.Column(db.Float, nullable=False)

class Approval(db.Model):
    __tablename__ = 'approvals'
    id = db.Column(db.Integer, primary_key=True)
    requisition_id = db.Column(db.Integer, db.ForeignKey('requisitions.id'), nullable=False)
    approver_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    approval_level = db.Column(db.Integer, nullable=False)
    role_name = db.Column(db.String(50), nullable=False)
    status = db.Column(db.String(20), default='Pending')
    comments = db.Column(db.Text)
    date_approved = db.Column(db.DateTime)

# ==================== HELPER FUNCTIONS ====================

def get_approval_sequence():
    return {
        1: {'role': 'Supervisor', 'next_level': 2, 'next_role': 'HOD'},
        2: {'role': 'HOD', 'next_level': 3, 'next_role': 'Procurement'},
        3: {'role': 'Procurement', 'next_level': 4, 'next_role': 'Finance'},
        4: {'role': 'Finance', 'next_level': 5, 'next_role': 'GM'},
        5: {'role': 'GM', 'next_level': None, 'next_role': 'Completed'}
    }

def get_role_for_level(level):
    sequence = get_approval_sequence()
    return sequence.get(level, {}).get('role', None)

def get_requisition_total(requisition):
    return sum(item.quantity * item.estimated_cost for item in requisition.items)

# ==================== AUTHENTICATION DECORATORS ====================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please log in to access this page.', 'warning')
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated_function

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if session.get('user_role') not in roles:
                flash('You do not have permission to access this page.', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

# ==================== ROUTES ====================

@app.route('/')
def login_page():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    departments = Department.query.order_by(Department.name).all()
    roles = ['Employee', 'Supervisor', 'HOD', 'HR', 'Procurement', 'Finance', 'GM']
    return render_template('login.html', roles=roles, departments=departments)

@app.route('/login', methods=['POST'])
def login():
    action = request.form.get('action')
    
    if action == 'login':
        email = request.form.get('email')
        password = request.form.get('password')
        
        user = User.query.filter_by(email=email, is_active=True).first()
        
        if user and check_password_hash(user.password, password):
            session['user_id'] = user.id
            session['user_name'] = user.name
            session['user_email'] = user.email
            session['user_role'] = user.role
            session['department_id'] = user.department_id
            if user.department_ref:
                session['department_name'] = user.department_ref.name
            else:
                session['department_name'] = None
            
            flash(f'Welcome back, {user.name}!', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid email or password.', 'danger')
            
    elif action == 'register':
        name = request.form.get('name')
        email = request.form.get('email')
        role = request.form.get('role')
        password = request.form.get('password')
        department_id = request.form.get('department_id') or None
        if department_id:
            department_id = int(department_id)
        
        existing_user = User.query.filter_by(email=email).first()
        if existing_user:
            flash('Email already registered. Please login instead.', 'warning')
            return redirect(url_for('login_page'))
        
        hashed_password = generate_password_hash(password, method='pbkdf2:sha256')
        
        new_user = User(
            name=name,
            email=email,
            password=hashed_password,
            role=role,
            department_id=department_id
        )
        
        db.session.add(new_user)
        db.session.commit()
        
        flash('Account created successfully! Please login.', 'success')
    
    return redirect(url_for('login_page'))

@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out.', 'info')
    return redirect(url_for('login_page'))

@app.route('/dashboard')
@login_required
def dashboard():
    user_role = session['user_role']
    user_id = session['user_id']
    
    if user_role == 'GM':
        requisitions = Requisition.query.filter(
            Requisition.current_approval_level == 5
        ).order_by(Requisition.date_created.desc()).all()
    elif user_role in ['Finance', 'Procurement', 'HOD', 'HR']:
        level_map = {'HR': 2, 'HOD': 2, 'Procurement': 3, 'Finance': 4}
        level = level_map.get(user_role)
        requisitions = Requisition.query.filter(
            Requisition.current_approval_level == level
        ).order_by(Requisition.date_created.desc()).all()
    elif user_role == 'Supervisor':
        department_id = session.get('department_id')
        if department_id:
            requisitions = Requisition.query.filter(
                Requisition.department_id == department_id,
                Requisition.current_approval_level == 1
            ).order_by(Requisition.date_created.desc()).all()
        else:
            requisitions = []
    else:
        requisitions = Requisition.query.filter_by(requestor_id=user_id).order_by(Requisition.date_created.desc()).all()
    
    for req in requisitions:
        req.current_approver_role = get_role_for_level(req.current_approval_level)
    
    return render_template('dashboard.html', requisitions=requisitions)

@app.route('/new_requisition', methods=['GET', 'POST'])
@login_required
@role_required('Employee')
def new_requisition():
    if request.method == 'POST':
        reason = request.form.get('reason')
        descriptions = request.form.getlist('description[]')
        quantities = request.form.getlist('quantity[]')
        costs = request.form.getlist('cost[]')
        
        if not descriptions or not descriptions[0]:
            flash('Please add at least one item to your requisition.', 'danger')
            return redirect(url_for('new_requisition'))
        
        user = User.query.get(session['user_id'])
        
        if not user.department_id:
            flash('Your account has not been assigned to a department. Please contact HR.', 'danger')
            return redirect(url_for('dashboard'))
        
        requisition = Requisition(
            requestor_id=session['user_id'],
            department_id=user.department_id,
            reason=reason,
            status='Pending Supervisor Approval',
            current_approval_level=1
        )
        db.session.add(requisition)
        db.session.flush()
        
        for desc, qty, cost in zip(descriptions, quantities, costs):
            if desc and qty and cost:
                item = RequisitionItem(
                    requisition_id=requisition.id,
                    description=desc,
                    quantity=int(qty),
                    estimated_cost=float(cost)
                )
                db.session.add(item)
        
        approval_sequence = get_approval_sequence()
        for level in range(1, 6):
            approval = Approval(
                requisition_id=requisition.id,
                approver_id=None,
                approval_level=level,
                role_name=approval_sequence[level]['role'],
                status='Pending'
            )
            db.session.add(approval)
        
        db.session.commit()
        
        flash(f'Requisition #{requisition.id} submitted successfully!', 'success')
        return redirect(url_for('dashboard'))
    
    return render_template('requisition.html', datetime_now=datetime.now().strftime('%B %d, %Y'))

@app.route('/handle_action/<int:req_id>/<string:action>')
@login_required
def handle_action(req_id, action):
    requisition = Requisition.query.get_or_404(req_id)
    user_role = session['user_role']
    user_id = session['user_id']
    current_level = requisition.current_approval_level
    required_role = get_role_for_level(current_level)
    
    if user_role != required_role:
        flash('You are not authorized to perform this action.', 'danger')
        return redirect(url_for('dashboard'))
    
    approval = Approval.query.filter_by(
        requisition_id=req_id,
        approval_level=current_level
    ).first()
    
    if approval:
        approval.status = 'Approved' if action == 'approve' else 'Rejected'
        approval.approver_id = user_id
        approval.comments = f"{'Approved' if action == 'approve' else 'Rejected'} by {session['user_name']}"
        approval.date_approved = datetime.utcnow()
    
    if action == 'reject':
        requisition.status = 'Rejected'
        db.session.commit()
        flash(f'Requisition #{requisition.id} has been REJECTED.', 'warning')
    else:
        next_level = get_approval_sequence().get(current_level, {}).get('next_level')
        
        if next_level is None:
            requisition.status = 'Approved'
            requisition.current_approval_level = 5
            db.session.commit()
            flash(f'Requisition #{requisition.id} has been FULLY APPROVED!', 'success')
        else:
            next_role = get_role_for_level(next_level)
            requisition.current_approval_level = next_level
            requisition.status = f'Pending {next_role} Approval'
            db.session.commit()
            flash(f'Requisition #{requisition.id} approved. Now pending {next_role} approval.', 'success')
    
    return redirect(url_for('dashboard'))

@app.route('/requisition_status/<int:req_id>')
@login_required
def requisition_status(req_id):
    requisition = Requisition.query.get_or_404(req_id)
    
    if requisition.requestor_id != session['user_id'] and session['user_role'] not in ['GM', 'Finance', 'Procurement', 'HOD', 'HR', 'Supervisor']:
        flash('You are not authorized to view this requisition.', 'danger')
        return redirect(url_for('dashboard'))
    
    approvals = Approval.query.filter_by(requisition_id=req_id).order_by(Approval.approval_level).all()
    total_value = get_requisition_total(requisition)
    
    return render_template('requisition_status.html', 
                          requisition=requisition, 
                          approvals=approvals, 
                          total_value=total_value,
                          get_role_for_level=get_role_for_level)

@app.route('/delete_requisition/<int:req_id>')
@login_required
def delete_requisition(req_id):
    requisition = Requisition.query.get_or_404(req_id)
    
    if requisition.requestor_id != session['user_id']:
        flash('You are not authorized to delete this requisition.', 'danger')
        return redirect(url_for('dashboard'))
    
    Approval.query.filter_by(requisition_id=req_id).delete()
    RequisitionItem.query.filter_by(requisition_id=req_id).delete()
    db.session.delete(requisition)
    db.session.commit()
    
    flash(f'Requisition #{req_id} has been deleted.', 'success')
    return redirect(url_for('dashboard'))

# ==================== ADMIN SETUP ====================

@app.route('/admin/setup', methods=['GET', 'POST'])
def admin_setup():
    if User.query.first():
        return "System already initialized. <a href='/'>Go to Login</a>"
    
    if request.method == 'POST':
        # Create departments
        depts = ['IT', 'Finance', 'HR', 'Operations', 'Sales', 'Procurement', 'Engineering', 'Marketing']
        dept_objects = {}
        
        for dept_name in depts:
            dept = Department(name=dept_name, description=f'{dept_name} Department')
            db.session.add(dept)
            db.session.flush()
            dept_objects[dept_name] = dept
        
        # Create GM
        gm = User(
            name='John Kamau',
            email='gm@saferpower.com',
            password=generate_password_hash('GM@2024', method='pbkdf2:sha256'),
            role='GM',
            department_id=None
        )
        db.session.add(gm)
        
        # Create HR
        hr = User(
            name='Sarah Wanjiku',
            email='hr@saferpower.com',
            password=generate_password_hash('HR@2024', method='pbkdf2:sha256'),
            role='HR',
            department_id=dept_objects['HR'].id
        )
        db.session.add(hr)
        
        # Create Procurement
        proc = User(
            name='James Mwangi',
            email='procurement@saferpower.com',
            password=generate_password_hash('Procurement@2024', method='pbkdf2:sha256'),
            role='Procurement',
            department_id=dept_objects['Procurement'].id
        )
        db.session.add(proc)
        
        # Create Finance
        fin = User(
            name='Grace Atieno',
            email='finance@saferpower.com',
            password=generate_password_hash('Finance@2024', method='pbkdf2:sha256'),
            role='Finance',
            department_id=dept_objects['Finance'].id
        )
        db.session.add(fin)
        
        # Create HODs
        for dept_name in depts:
            hod = User(
                name=f'{dept_name} HOD',
                email=f'hod.{dept_name.lower()}@saferpower.com',
                password=generate_password_hash('HOD@2024', method='pbkdf2:sha256'),
                role='HOD',
                department_id=dept_objects[dept_name].id
            )
            db.session.add(hod)
        
        # Create Supervisors
        for dept_name in depts:
            sup = User(
                name=f'{dept_name} Supervisor',
                email=f'sup.{dept_name.lower()}@saferpower.com',
                password=generate_password_hash('Supervisor@2024', method='pbkdf2:sha256'),
                role='Supervisor',
                department_id=dept_objects[dept_name].id
            )
            db.session.add(sup)
            db.session.flush()
            
            dept_sup = DepartmentSupervisor(
                department_id=dept_objects[dept_name].id,
                supervisor_id=sup.id
            )
            db.session.add(dept_sup)
        
        # Create Employees
        employees = [
            ('John Doe', 'john.doe@saferpower.com', 'IT'),
            ('Jane Smith', 'jane.smith@saferpower.com', 'Finance'),
            ('Alice Brown', 'alice.brown@saferpower.com', 'HR'),
        ]
        
        for name, email, dept in employees:
            emp = User(
                name=name,
                email=email,
                password=generate_password_hash('Employee@123', method='pbkdf2:sha256'),
                role='Employee',
                department_id=dept_objects[dept].id
            )
            db.session.add(emp)
        
        db.session.commit()
        
        return """
        <html>
        <body style="font-family: Arial; padding: 20px; text-align: center;">
            <h2 style="color: green;">✅ System Setup Complete!</h2>
            <h3>Login Credentials:</h3>
            <ul style="display: inline-block; text-align: left;">
                <li><strong>GM:</strong> gm@saferpower.com / GM@2024</li>
                <li><strong>HR:</strong> hr@saferpower.com / HR@2024</li>
                <li><strong>Finance:</strong> finance@saferpower.com / Finance@2024</li>
                <li><strong>Procurement:</strong> procurement@saferpower.com / Procurement@2024</li>
                <li><strong>Employee:</strong> john.doe@saferpower.com / Employee@123</li>
            </ul>
            <br>
            <a href="/" style="background: #f0a500; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Go to Login</a>
        </body>
        </html>
        """
    
    return '''
    <html>
    <body style="font-family: Arial; padding: 20px; text-align: center;">
        <h2>Initialize System</h2>
        <p>This will create departments, users, and roles.</p>
        <form method="POST">
            <button type="submit" style="background: green; color: white; padding: 10px 20px; font-size: 16px;">Initialize System</button>
        </form>
    </body>
    </html>
    '''

@app.route('/admin/reset')
def admin_reset():
    if os.path.exists('requisition_system.db'):
        os.remove('requisition_system.db')
    db.create_all()
    return "Database reset. <a href='/admin/setup'>Go to Setup</a>"

# ==================== HEALTH CHECK ====================

@app.route('/health')
def health():
    return {"status": "ok", "message": "Safer Power System Running"}

# ==================== RUN APP ====================

with app.app_context():
    db.create_all()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
