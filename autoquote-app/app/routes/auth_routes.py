from flask import Blueprint, render_template, redirect, request, session, url_for

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        senha = request.form.get('senha')
        if email == 'admin@autoquote.com' and senha == '123456':
            session['user'] = email
            return redirect(url_for('main.dashboard'))
        return render_template('login.html', error='Credenciais inválidas.')
    return render_template('login.html')

@auth_bp.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('auth.login'))
