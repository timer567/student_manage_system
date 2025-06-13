from flask import Flask, Response
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
import io
import pandas as pd
from flask import flash, redirect, url_for, request, session
from datetime import datetime
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from new1.app import html_footer

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///students.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)


# 数据库模型
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), nullable=False, default='teacher')

class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    sno = db.Column(db.String(20), unique=True, nullable=False)
    name = db.Column(db.String(80), nullable=False)
    score1 = db.Column(db.Float, nullable=False)
    score2 = db.Column(db.Float, nullable=False)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.now)

class Course(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), nullable=False)
    credit = db.Column(db.Float, nullable=False)

# 初始化数据库
with app.app_context():
    db.create_all()
    # 创建默认管理员账户
    if not User.query.filter_by(username='admin').first():
        admin = User(
            username='admin',
            password=generate_password_hash('admin123'),
            role='admin'
        )
        db.session.add(admin)
        db.session.commit()

def render_css():
    return """
    /* Bootstrap CSS */
    @import url('https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css');

    body {
        font-family: 'Microsoft YaHei', Arial, sans-serif;
        background: #f8f9fa;
        padding-top: 60px;
    }

    .navbar {
        background-color: #343a40;
        box-shadow: 0 2px 4px rgba(0,0,0,.1);
    }

    .card {
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,.1);
    }

    .btn-custom {
        margin: 5px;
        transition: all 0.3s ease;
    }

    .btn-custom:hover {
        transform: translateY(-2px);
        box-shadow: 0 2px 4px rgba(0,0,0,.2);
    }

    .table {
        background: white;
        border-radius: 5px;
        overflow: hidden;
    }

    .chart-container {
        background: white;
        padding: 20px;
        border-radius: 5px;
        margin-bottom: 20px;
    }

    .stats-card {
        transition: all 0.3s ease;
    }

    .stats-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 4px 8px rgba(0,0,0,.2);
    }
    """

def html_header(title="学生成绩管理系统"):
    # 根据session状态生成登录/退出按钮
    auth_button = """
        <li class="nav-item"><a class="nav-link" href="/logout"><i class="fas fa-sign-out-alt"></i> 退出</a></li>
    """ if session.get('logged_in') else """
        <li class="nav-item"><a class="nav-link" href="/login"><i class="fas fa-sign-in-alt"></i> 登录</a></li>
    """

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        <title>{title}</title>
        <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
        <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/css/all.min.css" rel="stylesheet">
        <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <style>{render_css()}</style>
    </head>
    <body>
        <nav class="navbar navbar-expand-lg navbar-dark bg-dark fixed-top">
            <div class="container">
                <a class="navbar-brand" href="/">学生成绩管理系统</a>
                <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNav">
                    <span class="navbar-toggler-icon"></span>
                </button>
                <div class="collapse navbar-collapse" id="navbarNav">
                    <ul class="navbar-nav mr-auto">
                        <li class="nav-item"><a class="nav-link" href="/"><i class="fas fa-home"></i> 首页</a></li>
                        <li class="nav-item"><a class="nav-link" href="/add"><i class="fas fa-user-plus"></i> 添加学生</a></li>
                        <li class="nav-item"><a class="nav-link" href="/list"><i class="fas fa-list"></i> 学生列表</a></li>
                        <li class="nav-item"><a class="nav-link" href="/stats"><i class="fas fa-chart-bar"></i> 统计分析</a></li>
                        <li class="nav-item dropdown">
                            <a class="nav-link dropdown-toggle" href="#" id="navbarDropdown" role="button" data-toggle="dropdown">
                                <i class="fas fa-cog"></i> 更多功能
                            </a>
                            <div class="dropdown-menu">
                                <a class="dropdown-item" href="/sort_save"><i class="fas fa-sort-amount-down"></i> 排序保存</a>
                                <a class="dropdown-item" href="/import_export"><i class="fas fa-file-import"></i> 导入导出</a>
                                <a class="dropdown-item" href="/backup"><i class="fas fa-database"></i> 数据备份</a>
                            </div>
                        </li>
                    </ul>
                    <ul class="navbar-nav">
                        {auth_button}
                    </ul>
                </div>
            </div>
        </nav>
        <div class="container mt-4">
            <h2 class="mb-4">{title}</h2>
    """

# 登录相关功能
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()

        if user and check_password_hash(user.password, password):
            session['logged_in'] = True
            session['username'] = username
            session['role'] = user.role
            flash('登录成功！', 'success')
            return redirect(url_for('index'))
        flash('用户名或密码错误！', 'danger')

    return f"""
    {html_header("用户登录")}
    <div class="row justify-content-center">
        <div class="col-md-6">
            <div class="card">
                <div class="card-body">
                    <form method="post">
                        <div class="form-group">
                            <label>用户名</label>
                            <input type="text" class="form-control" name="username" required>
                        </div>
                        <div class="form-group">
                            <label>密码</label>
                            <input type="password" class="form-control" name="password" required>
                        </div>
                        <button type="submit" class="btn btn-primary btn-block">登录</button>
                    </form>
                </div>
            </div>
        </div>
    </div>
    {html_footer()}
    """

@app.route('/logout')
def logout():
    session.clear()
    flash('已退出登录！', 'info')
    return redirect(url_for('login'))

# 首页
@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    # 获取基础统计数据
    total_students = Student.query.count()
    avg_score = db.session.query(db.func.avg(Student.score1 + Student.score2)).scalar() or 0
    pass_count = Student.query.filter(Student.score1 >= 60, Student.score2 >= 60).count()

    return f"""
    {html_header()}
    <div class="row">
        <div class="col-md-4">
            <div class="card stats-card bg-primary text-white">
                <div class="card-body">
                    <h5 class="card-title"><i class="fas fa-users"></i> 学生总数</h5>
                    <h2>{total_students}</h2>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card stats-card bg-success text-white">
                <div class="card-body">
                    <h5 class="card-title"><i class="fas fa-chart-line"></i> 平均总分</h5>
                    <h2>{avg_score:.1f}</h2>
                </div>
            </div>
        </div>
        <div class="col-md-4">
            <div class="card stats-card bg-info text-white">
                <div class="card-body">
                    <h5 class="card-title"><i class="fas fa-check-circle"></i> 及格人数</h5>
                    <h2>{pass_count}</h2>
                </div>
            </div>
        </div>
    </div>

    <div class="row mt-4">
        <div class="col-md-6">
            <div class="card">
                <div class="card-body">
                    <h5 class="card-title">快速操作</h5>
                    <div class="btn-group-vertical w-100">
                        <a href="/add" class="btn btn-outline-primary mb-2"><i class="fas fa-user-plus"></i> 添加新学生</a>
                        <a href="/list" class="btn btn-outline-secondary mb-2"><i class="fas fa-list"></i> 查看学生列表</a>
                        <a href="/stats" class="btn btn-outline-info mb-2"><i class="fas fa-chart-bar"></i> 查看统计分析</a>
                        <a href="/import_export" class="btn btn-outline-success"><i class="fas fa-file-import"></i> 导入/导出数据</a>
                    </div>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="card">
                <div class="card-body">
                    <h5 class="card-title">最近添加的学生</h5>
                    <div class="table-responsive">
                        <table class="table table-sm">
                            <thead>
                                <tr>
                                    <th>学号</th>
                                    <th>姓名</th>
                                    <th>添加时间</th>
                                </tr>
                            </thead>
                            <tbody>
                                {generate_recent_students_table()}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>
    {html_footer()}
    """

def generate_recent_students_table():
    recent_students = Student.query.order_by(Student.created_at.desc()).limit(5).all()
    rows = ""
    for student in recent_students:
        rows += f"""
        <tr>
            <td>{student.sno}</td>
            <td>{student.name}</td>
            <td>{student.created_at.strftime('%Y-%m-%d %H:%M')}</td>
        </tr>
        """
    return rows

# 添加学生
@app.route('/add', methods=['GET', 'POST'])
def add_student():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    if request.method == 'POST':
        sno = request.form.get('sno', '').strip()
        name = request.form.get('name', '').strip()
        try:
            score1 = float(request.form.get('score1', ''))
            score2 = float(request.form.get('score2', ''))
        except ValueError:
            flash('成绩必须是数字！', 'danger')
            return redirect(url_for('add_student'))

        if not sno or not name:
            flash('学号和姓名不能为空！', 'danger')
            return redirect(url_for('add_student'))

        # 检查学号是否已存在
        if Student.query.filter_by(sno=sno).first():
            flash('该学号已存在！', 'danger')
            return redirect(url_for('add_student'))

        # 创建新学生记录
        new_student = Student(sno=sno, name=name, score1=score1, score2=score2)
        db.session.add(new_student)
        db.session.commit()

        flash('学生添加成功！', 'success')
        return redirect(url_for('list_students'))

    return f"""
    {html_header("添加学生")}
    <div class="row justify-content-center">
        <div class="col-md-8">
            <div class="card">
                <div class="card-body">
                    <form method="post" class="needs-validation" novalidate>
                        <div class="form-group">
                            <label>学号</label>
                            <input type="text" class="form-control" name="sno" required>
                        </div>
                        <div class="form-group">
                            <label>姓名</label>
                            <input type="text" class="form-control" name="name" required>
                        </div>
                        <div class="form-group">
                            <label>课程1成绩</label>
                            <input type="number" class="form-control" name="score1" step="0.1" required>
                        </div>
                        <div class="form-group">
                            <label>课程2成绩</label>
                            <input type="number" class="form-control" name="score2" step="0.1" required>
                        </div>
                        <button type="submit" class="btn btn-primary"><i class="fas fa-save"></i> 保存</button>
                        <a href="/list" class="btn btn-secondary"><i class="fas fa-times"></i> 取消</a>
                    </form>
                </div>
            </div>
        </div>
    </div>
    {html_footer()}
    """

# 学生列表
@app.route('/list')
@app.route('/list/<sort_by>')
def list_students(sort_by='sno'):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    # 获取排序方向
    sort_direction = request.args.get('direction', 'asc')

    # 定义排序规则
    sort_rules = {
        'sno': Student.sno,
        'name': Student.name,
        'score1': Student.score1,
        'score2': Student.score2,
        'total': db.func.sum(Student.score1 + Student.score2)
    }

    # 获取要排序的列
    sort_column = sort_rules.get(sort_by, Student.sno)

    # 应用排序
    if sort_by == 'total':
        # 对总分特殊处理
        students = Student.query.all()
        students = sorted(students,
                          key=lambda x: (x.score1 + x.score2),
                          reverse=(sort_direction == 'desc'))
    else:
        # 其他列的排序
        if sort_direction == 'desc':
            students = Student.query.order_by(sort_column.desc()).all()
        else:
            students = Student.query.order_by(sort_column.asc()).all()

    # 生成排序链接
    def get_sort_link(column):
        new_direction = 'desc' if sort_by == column and sort_direction == 'asc' else 'asc'
        return f'/list/{column}?direction={new_direction}'

    # 生成排序图标
    def get_sort_icon(column):
        if sort_by == column:
            return 'fa-sort-up' if sort_direction == 'asc' else 'fa-sort-down'
        return 'fa-sort'

    rows = ""
    for student in students:
        total = student.score1 + student.score2
        created_time = student.created_at.strftime('%Y-%m-%d %H:%M:%S') if student.created_at else 'Unknown'
        rows += f"""
        <tr>
            <td>{student.sno}</td>
            <td>{student.name}</td>
            <td>{student.score1}</td>
            <td>{student.score2}</td>
            <td>{total}</td>
            <td>{created_time}</td>
            <td>
                <div class="btn-group">
                    <a href="/edit/{student.sno}" class="btn btn-sm btn-info">
                        <i class="fas fa-edit"></i> 编辑
                    </a>
                    <a href="/delete/{student.sno}" class="btn btn-sm btn-danger" 
                       onclick="return confirm('确认删除该学生吗？')">
                        <i class="fas fa-trash"></i> 删除
                    </a>
                </div>
            </td>
        </tr>
        """

    return f"""
    {html_header("学生列表")}
    <div class="card">
        <div class="card-body">
            <div class="table-responsive">
                <table class="table table-striped table-hover">
                    <thead class="thead-dark">
                        <tr>
                            <th>
                                <a href="{get_sort_link('sno')}" class="text-white" style="text-decoration: none">
                                    学号 <i class="fas {get_sort_icon('sno')}"></i>
                                </a>
                            </th>
                            <th>
                                <a href="{get_sort_link('name')}" class="text-white" style="text-decoration: none">
                                    姓名 <i class="fas {get_sort_icon('name')}"></i>
                                </a>
                            </th>
                            <th>
                                <a href="{get_sort_link('score1')}" class="text-white" style="text-decoration: none">
                                    课程1成绩 <i class="fas {get_sort_icon('score1')}"></i>
                                </a>
                            </th>
                            <th>
                                <a href="{get_sort_link('score2')}" class="text-white" style="text-decoration: none">
                                    课程2成绩 <i class="fas {get_sort_icon('score2')}"></i>
                                </a>
                            </th>
                            <th>
                                <a href="{get_sort_link('total')}" class="text-white" style="text-decoration: none">
                                    总成绩 <i class="fas {get_sort_icon('total')}"></i>
                                </a>
                            </th>
                            <th>录入时间</th>
                            <th>操作</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <style>
    th a {{
        display: block;
    }}

    th a:hover {{
        color: #fff;
    }}

    .fas.fa-sort,
    .fas.fa-sort-up,
    .fas.fa-sort-down {{
        margin-left: 5px;
    }}
    </style>
    {html_footer()}
    """

# 编辑学生
@app.route('/edit/<sno>', methods=['GET', 'POST'])
def edit_student(sno):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    student = Student.query.filter_by(sno=sno).first()
    if not student:
        flash('未找到该学生！', 'danger')
        return redirect(url_for('list_students'))

    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        try:
            score1 = float(request.form.get('score1', ''))
            score2 = float(request.form.get('score2', ''))
        except ValueError:
            flash('成绩必须是数字！', 'danger')
            return redirect(url_for('edit_student', sno=sno))

        if not name:
            flash('姓名不能为空！', 'danger')
            return redirect(url_for('edit_student', sno=sno))

        student.name = name
        student.score1 = score1
        student.score2 = score2
        db.session.commit()

        flash('修改成功！', 'success')
        return redirect(url_for('list_students'))

    return f"""
    {html_header(f"编辑学生 - {student.name}")}
    <div class="row justify-content-center">
        <div class="col-md-8">
            <div class="card">
                <div class="card-body">
                    <form method="post">
                        <div class="form-group">
                            <label>学号</label>
                            <input type="text" class="form-control" value="{student.sno}" readonly>
                        </div>
                        <div class="form-group">
                            <label>姓名</label>
                            <input type="text" class="form-control" name="name" value="{student.name}" required>
                        </div>
                        <div class="form-group">
                            <label>课程1成绩</label>
                            <input type="number" class="form-control" name="score1" value="{student.score1}" step="0.1" required>
                        </div>
                        <div class="form-group">
                            <label>课程2成绩</label>
                            <input type="number" class="form-control" name="score2" value="{student.score2}" step="0.1" required>
                        </div>
                        <button type="submit" class="btn btn-primary"><i class="fas fa-save"></i> 保存修改</button>
                        <a href="/list" class="btn btn-secondary"><i class="fas fa-times"></i> 取消</a>
                    </form>
                </div>
            </div>
        </div>
    </div>
    {html_footer()}
    """

# 删除学生
@app.route('/delete/<sno>')
def delete_student(sno):
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    student = Student.query.filter_by(sno=sno).first()
    if student:
        db.session.delete(student)
        db.session.commit()
        flash('学生已删除！', 'success')
    else:
        flash('未找到该学生！', 'danger')
    return redirect(url_for('list_students'))

# 统计分析
@app.route('/stats')
def stats():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    students = Student.query.all()
    if not students:
        return f"{html_header('统计分析')}<p>当前没有学生数据。</p>{html_footer()}"

    total = len(students)

    # 计算统计数据
    total_scores = [s.score1 + s.score2 for s in students]
    avg_score = sum(total_scores) / len(total_scores)

    # 计算课程1的等级分布
    def get_grade_stats(scores, students, score_field):
        fail = []
        pass_list = []
        good = []
        excellent = []

        for i, score in enumerate(scores):
            student = students[i]
            student_info = {
                'sno': student.sno,
                'name': student.name,
                'score': score
            }
            if score < 60:
                fail.append(student_info)
            elif score < 70:
                pass_list.append(student_info)
            elif score < 85:
                good.append(student_info)
            else:
                excellent.append(student_info)

        return {
            'fail': fail,
            'fail_count': len(fail),
            'fail_rate': (len(fail) / total * 100) if total > 0 else 0,
            'pass': pass_list,
            'pass_count': len(pass_list),
            'pass_rate': (len(pass_list) / total * 100) if total > 0 else 0,
            'good': good,
            'good_count': len(good),
            'good_rate': (len(good) / total * 100) if total > 0 else 0,
            'excellent': excellent,
            'excellent_count': len(excellent),
            'excellent_rate': (len(excellent) / total * 100) if total > 0 else 0
        }

    score1_stats = get_grade_stats([s.score1 for s in students], students, 'score1')
    score2_stats = get_grade_stats([s.score2 for s in students], students, 'score2')

    # 计算课程平均分
    avg_score1 = sum(s.score1 for s in students) / total if total > 0 else 0
    avg_score2 = sum(s.score2 for s in students) / total if total > 0 else 0

    # 成绩分布统计
    score_ranges = ['0-59', '60-69', '70-84', '85-100']
    score1_dist = [score1_stats['fail_count'], score1_stats['pass_count'],
                   score1_stats['good_count'], score1_stats['excellent_count']]
    score2_dist = [score2_stats['fail_count'], score2_stats['pass_count'],
                   score2_stats['good_count'], score2_stats['excellent_count']]

    return f"""
    {html_header('统计分析')}
    <div class="container mt-4">
        <h2 class="text-center mb-4">成绩统计分析</h2>

        <div class="row">
            <div class="col-md-6">
                <div class="card mb-4">
                    <div class="card-header bg-primary text-white">
                        <h5 class="card-title mb-0">课程1成绩分布</h5>
                    </div>
                    <div class="card-body">
                        <div class="table-responsive">
                            <table class="table table-bordered">
                                <thead>
                                    <tr>
                                        <th>等级</th>
                                        <th>人数</th>
                                        <th>比例</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr class="table-danger grade-row" data-course="1" data-grade="fail">
                                        <td>不及格 (<60分)</td>
                                        <td>{score1_stats['fail_count']}人</td>
                                        <td>{score1_stats['fail_rate']:.1f}%</td>
                                    </tr>
                                    <tr class="table-warning grade-row" data-course="1" data-grade="pass">
                                        <td>及格 (60-69分)</td>
                                        <td>{score1_stats['pass_count']}人</td>
                                        <td>{score1_stats['pass_rate']:.1f}%</td>
                                    </tr>
                                    <tr class="table-info grade-row" data-course="1" data-grade="good">
                                        <td>良好 (70-84分)</td>
                                        <td>{score1_stats['good_count']}人</td>
                                        <td>{score1_stats['good_rate']:.1f}%</td>
                                    </tr>
                                    <tr class="table-success grade-row" data-course="1" data-grade="excellent">
                                        <td>优秀 (≥85分)</td>
                                        <td>{score1_stats['excellent_count']}人</td>
                                        <td>{score1_stats['excellent_rate']:.1f}%</td>
                                    </tr>
                                    <tr class="table-active">
                                        <td><strong>总计</strong></td>
                                        <td><strong>{total}人</strong></td>
                                        <td><strong>100%</strong></td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <div class="col-md-6">
                <div class="card mb-4">
                    <div class="card-header bg-success text-white">
                        <h5 class="card-title mb-0">课程2成绩分布</h5>
                    </div>
                    <div class="card-body">
                        <div class="table-responsive">
                            <table class="table table-bordered">
                                <thead>
                                    <tr>
                                        <th>等级</th>
                                        <th>人数</th>
                                        <th>比例</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr class="table-danger grade-row" data-course="2" data-grade="fail">
                                        <td>不及格 (<60分)</td>
                                        <td>{score2_stats['fail_count']}人</td>
                                        <td>{score2_stats['fail_rate']:.1f}%</td>
                                    </tr>
                                    <tr class="table-warning grade-row" data-course="2" data-grade="pass">
                                        <td>及格 (60-69分)</td>
                                        <td>{score2_stats['pass_count']}人</td>
                                        <td>{score2_stats['pass_rate']:.1f}%</td>
                                    </tr>
                                    <tr class="table-info grade-row" data-course="2" data-grade="good">
                                        <td>良好 (70-84分)</td>
                                        <td>{score2_stats['good_count']}人</td>
                                        <td>{score2_stats['good_rate']:.1f}%</td>
                                    </tr>
                                    <tr class="table-success grade-row" data-course="2" data-grade="excellent">
                                        <td>优秀 (≥85分)</td>
                                        <td>{score2_stats['excellent_count']}人</td>
                                        <td>{score2_stats['excellent_rate']:.1f}%</td>
                                    </tr>
                                    <tr class="table-active">
                                        <td><strong>总计</strong></td>
                                        <td><strong>{total}人</strong></td>
                                        <td><strong>100%</strong></td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- 学生详细信息模态框 -->
        <div class="modal fade" id="studentModal" tabindex="-1">
            <div class="modal-dialog modal-lg">
                <div class="modal-content">
                    <div class="modal-header">
                        <h5 class="modal-title">学生成绩详情</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                    </div>
                    <div class="modal-body">
                        <div class="table-responsive">
                            <table class="table table-bordered">
                                <thead>
                                    <tr>
                                        <th>学号</th>
                                        <th>姓名</th>
                                        <th>成绩</th>
                                    </tr>
                                </thead>
                                <tbody id="modalTableBody">
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="card mb-4">
            <div class="card-header bg-info text-white">
                <h5 class="card-title mb-0">总体统计</h5>
            </div>
            <div class="card-body">
                <div class="row">
                    <div class="col-md-6">
                        <ul class="list-group">
                            <li class="list-group-item d-flex justify-content-between align-items-center">
                                学生总数
                                <span class="badge bg-primary rounded-pill">{total}</span>
                            </li>
                            <li class="list-group-item d-flex justify-content-between align-items-center">
                                课程1平均分
                                <span class="badge bg-info rounded-pill">{avg_score1:.2f}</span>
                            </li>
                            <li class="list-group-item d-flex justify-content-between align-items-center">
                                课程2平均分
                                <span class="badge bg-info rounded-pill">{avg_score2:.2f}</span>
                            </li>
                        </ul>
                    </div>
                    <div class="col-md-6">
                        <div class="chart-container">
                            <canvas id="scoreDistChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
    // 图表数据
    var ctx = document.getElementById('scoreDistChart').getContext('2d');
    new Chart(ctx, {{
        type: 'bar',
        data: {{
            labels: {score_ranges},
            datasets: [{{
                label: '课程1成绩分布',
                data: {score1_dist},
                backgroundColor: 'rgba(54, 162, 235, 0.5)',
                borderColor: 'rgba(54, 162, 235, 1)',
                borderWidth: 1
            }},
            {{
                label: '课程2成绩分布',
                data: {score2_dist},
                backgroundColor: 'rgba(255, 99, 132, 0.5)',
                borderColor: 'rgba(255, 99, 132, 1)',
                borderWidth: 1
            }}]
        }},
        options: {{
            responsive: true,
            scales: {{
                y: {{
                    beginAtZero: true,
                    ticks: {{
                        stepSize: 1
                    }}
                }}
            }}
        }}
    }});

    // 存储成绩数据供JavaScript使用
    var gradeData = {{
        '1': {{}},
        '2': {{}}
    }};

    // 课程1的数据
    gradeData['1']['fail'] = {str(score1_stats['fail'])};
    gradeData['1']['pass'] = {str(score1_stats['pass'])};
    gradeData['1']['good'] = {str(score1_stats['good'])};
    gradeData['1']['excellent'] = {str(score1_stats['excellent'])};

    // 课程2的数据
    gradeData['2']['fail'] = {str(score2_stats['fail'])};
    gradeData['2']['pass'] = {str(score2_stats['pass'])};
    gradeData['2']['good'] = {str(score2_stats['good'])};
    gradeData['2']['excellent'] = {str(score2_stats['excellent'])};

    // 点击成绩等级行显示学生详情
    document.querySelectorAll('.grade-row').forEach(row => {{
        row.style.cursor = 'pointer';
        row.addEventListener('click', function() {{
            const course = this.getAttribute('data-course');
            const grade = this.getAttribute('data-grade');
            const students = gradeData[course][grade];

            // 清空并填充模态框表格
            const tableBody = document.getElementById('modalTableBody');
            tableBody.innerHTML = '';

            students.forEach(student => {{
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${{student.sno}}</td>
                    <td>${{student.name}}</td>
                    <td>${{student.score}}</td>
                `;
                tableBody.appendChild(row);
            }});

            // 显示模态框
            new bootstrap.Modal(document.getElementById('studentModal')).show();
        }});
    }});
    </script>
    {html_footer()}
    """

# 数据导入导出
@app.route('/import_export', methods=['GET', 'POST'])
def import_export():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    # 获取当前时间和用户
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    current_user = session.get('username', 'Anonymous')
    import_error = None

    if request.method == 'POST':
        if 'file' not in request.files:
            flash('没有选择文件！', 'danger')
            import_error = "导入错误：没有选择文件"
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('没有选择文件！', 'danger')
            import_error = "导入错误：没有选择文件"
            return redirect(request.url)

        # 检查文件大小（例如限制为5MB）
        file_content = file.read()
        if len(file_content) > 5 * 1024 * 1024:  # 5MB
            flash('文件太大，请上传5MB以内的文件！', 'danger')
            import_error = "导入错误：文件大小超过5MB限制"
            return redirect(request.url)
        file.seek(0)  # 重置文件指针

        # 检查文件扩展名
        allowed_extensions = {'.csv', '.xlsx'}
        file_ext = os.path.splitext(file.filename)[1].lower()

        if file_ext not in allowed_extensions:
            flash('请上传CSV或Excel(xlsx)格式的文件！', 'danger')
            import_error = "导入错误：文件格式不正确，仅支持CSV和Excel(xlsx)格式"
            return redirect(request.url)

        try:
            # 根据文件类型读取数据
            if file_ext == '.csv':
                # 尝试不同的编码方式和分隔符
                try:
                    # 读取CSV文件
                    df = pd.read_csv(file, encoding='utf-8-sig')  # 使用 utf-8-sig 来自动处理 BOM

                    # 确保列名正确（移除可能的BOM标记）
                    df.columns = df.columns.str.replace('\ufeff', '')

                    print("读取到的数据：")
                    print(df.head())
                    print(f"处理后的列名：{df.columns.tolist()}")

                except Exception as e:
                    # 如果上面的方法失败，尝试其他编码
                    try:
                        file.seek(0)
                        df = pd.read_csv(file, encoding='gbk')
                        df.columns = df.columns.str.replace('\ufeff', '')
                    except Exception:
                        try:
                            file.seek(0)
                            df = pd.read_csv(file, encoding='gb2312')
                            df.columns = df.columns.str.replace('\ufeff', '')
                        except Exception:
                            raise ValueError("无法正确读取CSV文件，请检查文件格式和编码")
            else:  # .xlsx
                df = pd.read_excel(file)

            print("最终读取到的数据：")
            print(df.head())
            print(f"列名：{df.columns.tolist()}")

            # 验证必要的列是否存在
            required_columns = ['学号', '姓名', '课程1成绩', '课程2成绩']
            missing_cols = [col for col in required_columns if col not in df.columns]

            if missing_cols:
                error_msg = f'文件格式错误！文件缺少以下必需列：{", ".join(missing_cols)}\n'
                error_msg += f'当前文件的列名：{", ".join(df.columns.tolist())}\n'
                error_msg += '请确保文件第一行包含以下列名：学号、姓名、课程1成绩、课程2成绩'
                flash(error_msg, 'danger')
                import_error = error_msg
                return redirect(request.url)

            if df.empty:
                flash('文件中没有数据！', 'danger')
                import_error = "导入错误：文件中没有数据"
                return redirect(request.url)

            # 数据验证
            error_rows = []
            success_count = 0

            for index, row in df.iterrows():
                try:
                    # 验证数据格式
                    sno = str(row['学号']).strip()
                    name = str(row['姓名']).strip()

                    try:
                        score1 = float(row['课程1成绩'])
                        score2 = float(row['课程2成绩'])
                    except ValueError:
                        raise ValueError("成绩必须为数字")

                    # 验证数据有效性
                    if not sno or not name:
                        raise ValueError("学号或姓名不能为空")
                    if not (0 <= score1 <= 100) or not (0 <= score2 <= 100):
                        raise ValueError("成绩必须在0-100之间")

                    # 检查学号是否已存在
                    existing_student = Student.query.filter_by(sno=sno).first()
                    if existing_student:
                        # 更新现有记录
                        existing_student.name = name
                        existing_student.score1 = score1
                        existing_student.score2 = score2
                    else:
                        # 创建新记录
                        student = Student(
                            sno=sno,
                            name=name,
                            score1=score1,
                            score2=score2
                        )
                        db.session.add(student)

                    success_count += 1
                except Exception as e:
                    error_msg = f"第{index + 2}行: {str(e)}"
                    error_rows.append(error_msg)

            if error_rows:
                # 如果有错误，回滚事务
                db.session.rollback()
                error_message = "导入过程中发现以下错误：\n" + "\n".join(error_rows)
                flash(error_message, 'danger')
                import_error = error_message
            else:
                try:
                    # 提交事务
                    db.session.commit()
                    flash(f'成功导入{success_count}条数据！', 'success')
                    return redirect(url_for('list_students'))
                except Exception as e:
                    db.session.rollback()
                    error_message = f'保存数据时出错：{str(e)}'
                    flash(error_message, 'danger')
                    import_error = error_message

        except Exception as e:
            error_message = f'读取文件失败：{str(e)}'
            flash(error_message, 'danger')
            import_error = error_message
            print(f"错误详情：{str(e)}")  # 打印详细错误信息
        return redirect(request.url)

    return f"""
    {html_header('数据导入导出')}
    <div class="container">
        <div class="row justify-content-center">
            <div class="col-md-6">
                <div class="card mb-4">
                    <div class="card-header bg-primary text-white">
                        <h5 class="card-title mb-0">导入数据</h5>
                    </div>
                    <div class="card-body">
                        {f'<div class="alert alert-danger">{import_error}</div>' if import_error else ''}
                        <form method="post" enctype="multipart/form-data" class="needs-validation" novalidate>
                            <div class="mb-3">
                                <label class="form-label">选择文件</label>
                                <input type="file" class="form-control" name="file" accept=".csv,.xlsx" required>
                                <div class="invalid-feedback">
                                    请选择一个文件
                                </div>
                                <small class="form-text text-muted">
                                    支持CSV和Excel(xlsx)格式文件（文件大小限制5MB）
                                </small>
                            </div>
                            <button type="submit" class="btn btn-primary">
                                <i class="fas fa-file-import"></i> 导入
                            </button>
                        </form>
                    </div>
                </div>
            </div>

            <div class="col-md-6">
                <div class="card mb-4">
                    <div class="card-header bg-success text-white">
                        <h5 class="card-title mb-0">导出数据</h5>
                    </div>
                    <div class="card-body">
                        <p class="card-text">选择导出格式：</p>
                        <div class="d-grid gap-2">
                            <a href="/export_csv" class="btn btn-success">
                                <i class="fas fa-file-export"></i> 导出为CSV
                            </a>
                            <a href="/export_pdf" class="btn btn-danger">
                                <i class="fas fa-file-pdf"></i> 导出为PDF
                            </a>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div class="card mb-4">
            <div class="card-header bg-info text-white">
                <h5 class="card-title mb-0">导入说明</h5>
            </div>
            <div class="card-body">
                <div class="alert alert-info" role="alert">
                    <strong>系统信息：</strong><br>
                    当前时间：{current_time}<br>
                    当前用户：{current_user}
                </div>
                <h6>文件要求：</h6>
                <ul>
                    <li>支持的文件格式：CSV和Excel(xlsx)</li>
                    <li>文件必须包含以下列：学号、姓名、课程1成绩、课程2成绩</li>
                    <li>成绩必须为0-100之间的数字</li>
                    <li>学号和姓名不能为空</li>
                    <li>文件大小不能超过5MB</li>
                </ul>
                <h6>注意事项：</h6>
                <ul>
                    <li>如果导入的学号已存在，将更新该学生的信息</li>
                    <li>建议先导出一份CSV文件作为模板参考</li>
                </ul>
            </div>
        </div>
    </div>
    {html_footer()}
    """

# 导出CSV
@app.route('/export_csv')
def export_csv():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    students = Student.query.all()

    # 创建一个StringIO对象来写入CSV数据
    output = io.StringIO()
    output.write('\ufeff')  # 添加BOM标记，解决Excel打开中文乱码问题

    # 写入表头
    headers = ['学号', '姓名', '课程1成绩', '课程2成绩', '总成绩', '录入时间']
    output.write(','.join(headers) + '\n')

    # 写入数据
    for student in students:
        total = student.score1 + student.score2
        created_time = student.created_at.strftime('%Y-%m-%d %H:%M:%S') if student.created_at else ''
        row = [
            str(student.sno),
            student.name,
            str(student.score1),
            str(student.score2),
            str(total),
            created_time
        ]
        output.write(','.join(row) + '\n')

    # 生成响应
    output.seek(0)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return Response(
        output.getvalue().encode('utf-8-sig'),
        mimetype='text/csv',
        headers={
            "Content-Disposition": f"attachment; filename=student_scores_{timestamp}.csv",
            "Content-Type": "text/csv; charset=utf-8-sig"
        }
    )

# 导出PDF
@app.route('/export_pdf')
def export_pdf():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    students = Student.query.all()

    # 创建PDF
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)

    # 设置中文字体
    pdfmetrics.registerFont(TTFont('SimSun', 'simsun.ttc'))
    p.setFont('SimSun', 12)

    # 添加标题
    p.drawString(250, 750, "学生成绩表")

    # 添加表头
    headers = ['学号', '姓名', '课程1成绩', '课程2成绩', '总成绩']
    x_positions = [50, 150, 250, 350, 450]
    for header, x in zip(headers, x_positions):
        p.drawString(x, 700, header)

    # 添加数据
    y = 670
    for student in students:
        if y < 50:  # 如果页面空间不足，新建一页
            p.showPage()
            y = 750

        total = student.score1 + student.score2
        data = [student.sno, student.name, str(student.score1),
                str(student.score2), str(total)]

        for value, x in zip(data, x_positions):
            p.drawString(x, y, value)

        y -= 20

    p.save()

    buffer.seek(0)
    return Response(
        buffer.getvalue(),
        mimetype='application/pdf',
        headers={
            'Content-Disposition': 'attachment; filename=student_scores.pdf',
            'Content-Type': 'application/pdf'
        }
    )

@app.route('/sort_save')
def sort_save():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    students = Student.query.all()
    # 按总分排序（由高到低）
    sorted_students = sorted(students, key=lambda x: (x.score1 + x.score2), reverse=True)

    # 计算各类比例
    total = len(students)
    if total == 0:
        return f"{html_header('排序与统计')}<p>当前没有学生数据。</p>{html_footer()}"

    fail_count = sum(1 for s in students if s.score1 < 60 or s.score2 < 60)
    pass_count = total - fail_count
    good_count = sum(1 for s in students if 150 <= (s.score1 + s.score2) < 170)
    excellent_count = sum(1 for s in students if (s.score1 + s.score2) >= 170)

    # 计算比例
    fail_ratio = fail_count / total * 100
    pass_ratio = pass_count / total * 100
    good_ratio = good_count / total * 100
    excellent_ratio = excellent_count / total * 100

    # 保存到文件
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'student_scores_{timestamp}.txt'

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("学生成绩排名表\n")
        f.write("=" * 50 + "\n")
        f.write("排名\t学号\t姓名\t课程1\t课程2\t总分\n")
        f.write("-" * 50 + "\n")

        for idx, student in enumerate(sorted_students, 1):
            total_score = student.score1 + student.score2
            f.write(f"{idx}\t{student.sno}\t{student.name}\t{student.score1}\t{student.score2}\t{total_score}\n")

        f.write("\n\n成绩分析\n")
        f.write("=" * 50 + "\n")
        f.write(f"不及格比例：{fail_ratio:.2f}%\n")
        f.write(f"及格比例：{pass_ratio:.2f}%\n")
        f.write(f"良好比例（总分150-169）：{good_ratio:.2f}%\n")
        f.write(f"优秀比例（总分≥170）：{excellent_ratio:.2f}%\n")

    # 生成网页显示内容
    rows = ""
    for idx, student in enumerate(sorted_students, 1):
        total_score = student.score1 + student.score2
        rows += f"""
        <tr>
            <td>{idx}</td>
            <td>{student.sno}</td>
            <td>{student.name}</td>
            <td>{student.score1}</td>
            <td>{student.score2}</td>
            <td>{total_score}</td>
        </tr>
        """

    return f"""
    {html_header("排序与统计")}
    <div class="alert alert-success">
        成绩已保存到文件：{filename}
    </div>

    <div class="row mb-4">
        <div class="col-md-6">
            <div class="card">
                <div class="card-body">
                    <h5 class="card-title">成绩比例分析</h5>
                    <ul class="list-group">
                        <li class="list-group-item d-flex justify-content-between align-items-center">
                            不及格比例
                            <span class="badge badge-danger badge-pill">{fail_ratio:.2f}%</span>
                        </li>
                        <li class="list-group-item d-flex justify-content-between align-items-center">
                            及格比例
                            <span class="badge badge-success badge-pill">{pass_ratio:.2f}%</span>
                        </li>
                        <li class="list-group-item d-flex justify-content-between align-items-center">
                            良好比例
                            <span class="badge badge-info badge-pill">{good_ratio:.2f}%</span>
                        </li>
                        <li class="list-group-item d-flex justify-content-between align-items-center">
                            优秀比例
                            <span class="badge badge-warning badge-pill">{excellent_ratio:.2f}%</span>
                        </li>
                    </ul>
                </div>
            </div>
        </div>
    </div>

    <div class="card">
        <div class="card-body">
            <h5 class="card-title">学生成绩排名</h5>
            <div class="table-responsive">
                <table class="table table-striped">
                    <thead class="thead-dark">
                        <tr>
                            <th>排名</th>
                            <th>学号</th>
                            <th>姓名</th>
                            <th>课程1成绩</th>
                            <th>课程2成绩</th>
                            <th>总分</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    {html_footer()}
    """

# 数据备份
@app.route('/backup')
def backup():
    if not session.get('logged_in'):
        return redirect(url_for('login'))

    students = Student.query.all()
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'backup_{timestamp}.txt'

    with open(filename, 'w', encoding='utf-8') as f:
        f.write("学生成绩管理系统 - 数据备份\n")
        f.write("=" * 50 + "\n")
        f.write(f"备份时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("-" * 50 + "\n\n")

        f.write("学号\t姓名\t课程1\t课程2\t总分\t录入时间\n")
        f.write("-" * 50 + "\n")

        for student in students:
            total = student.score1 + student.score2
            created_time = student.created_at.strftime('%Y-%m-%d %H:%M:%S')
            f.write(f"{student.sno}\t{student.name}\t{student.score1}\t{student.score2}\t{total}\t{created_time}\n")

    return f"""
    {html_header("数据备份")}
    <div class="alert alert-success">
        <i class="fas fa-check-circle"></i> 数据已成功备份到文件：{filename}
    </div>
    <div class="card">
        <div class="card-body">
            <h5 class="card-title">备份信息</h5>
            <ul class="list-group">
                <li class="list-group-item">
                    <strong>备份文件名：</strong> {filename}
                </li>
                <li class="list-group-item">
                    <strong>备份时间：</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                </li>
                <li class="list-group-item">
                    <strong>记录总数：</strong> {len(students)}
                </li>
            </ul>
        </div>
    </div>
    {html_footer()}
    """

def html_footer():
    return """
        </div>
        <footer class="footer mt-auto py-3 bg-light">
            <div class="container text-center">
                <span class="text-muted">学生成绩管理系统 &copy; 2025</span>
            </div>
        </footer>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <script>
        // 添加Bootstrap表单验证
        (function() {
            'use strict';
            window.addEventListener('load', function() {
                var forms = document.getElementsByClassName('needs-validation');
                var validation = Array.prototype.filter.call(forms, function(form) {
                    form.addEventListener('submit', function(event) {
                        if (form.checkValidity() === false) {
                            event.preventDefault();
                            event.stopPropagation();
                        }
                        form.classList.add('was-validated');
                    }, false);
                });
            }, false);
        })();

        // 添加Flash消息自动消失
        setTimeout(function() {
            $('.alert').fadeOut('slow');
        }, 3000);
        </script>
    </body>
    </html>
    """

if __name__ == '__main__':
    with app.app_context():
        db.create_all()  # 确保数据库表已创建
    app.run(debug=True)