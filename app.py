from flask import Flask, redirect, url_for
from config import Config
from datetime import timedelta


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=7)

    from routes.superadmin import sa_bp
    from routes.agency     import agency_bp
    from routes.public     import public_bp

    app.register_blueprint(sa_bp)
    app.register_blueprint(agency_bp)
    app.register_blueprint(public_bp)

    @app.route('/')
    def index():
        return redirect(url_for('agency.login'))

    # Preload doc templates
    from utils.doc_engine import preload_templates
    with app.app_context():
        preload_templates()

    return app
