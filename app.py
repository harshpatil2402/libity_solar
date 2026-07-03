from flask import Flask, redirect
from config import Config
from datetime import timedelta
from werkzeug.middleware.proxy_fix import ProxyFix # <--- Added this line


def create_app():
    app = Flask(__name__)
    
    # Tells Flask to trust Cloudflare's HTTPS headers so redirects work perfectly
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1) # <--- Added this line
    
    app.config.from_object(Config)
    app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=7)

    from routes.superadmin import sa_bp
    from routes.agency     import agency_bp
    from routes.public     import public_bp

    app.register_blueprint(sa_bp)
    app.register_blueprint(agency_bp)
    app.register_blueprint(public_bp)

    # Use a relative redirect to prevent Flask from appending internal container ports
    @app.route('/')
    def index():
        return redirect('/agency/login')

    # Preload doc templates
    from utils.doc_engine import preload_templates
    with app.app_context():
        preload_templates()

    return app