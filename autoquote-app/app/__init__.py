from flask import Flask
from flask_session import Session

def create_app():
    app = Flask(__name__, static_folder="static", template_folder="templates")
    app.config.from_mapping({
        "SECRET_KEY": "autoquote-dev-key",
        "SESSION_TYPE": "filesystem",
    })
    Session(app)

    from app.routes.auth_routes import auth_bp
    from app.routes.main_routes import main_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)

    return app
