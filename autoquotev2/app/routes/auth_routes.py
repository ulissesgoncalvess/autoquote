from flask import Blueprint, render_template, request, redirect, url_for, session, flash

auth_bp = Blueprint('auth', __name__)

# Usuário fixo para teste
USUARIO_TESTE = {
    'email': 'admin@autoquote.com',
    'senha': '123456',
    'nome': 'Administrador AutoQuote'
}

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form.get('email')
        senha = request.form.get('senha')
        
        if email == USUARIO_TESTE['email'] and senha == USUARIO_TESTE['senha']:
            session['user'] = USUARIO_TESTE
            flash('Login realizado com sucesso!', 'success')
            return redirect(url_for('main.dashboard'))
        else:
            flash('Email ou senha incorretos!', 'error')
    
    return render_template('login.html')

@auth_bp.route('/logout')
def logout():
    session.pop('user', None)
    flash('Logout realizado com sucesso!', 'info')
    return redirect(url_for('auth.login'))