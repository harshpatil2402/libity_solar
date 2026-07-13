from flask import Flask, redirect, request, jsonify
from config import Config
from datetime import timedelta
from werkzeug.exceptions import HTTPException
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

    # Any route the frontend calls via fetch/JSON must never fall through to
    # Flask's default HTML error page — that breaks res.json() client-side
    # with a cryptic "Unexpected token '<'" error instead of a usable message.
    @app.errorhandler(Exception)
    def handle_exception(e):
        wants_json = request.path.startswith('/agency/api/') or (request.method == 'POST' and request.is_json)
        if not wants_json:
            if isinstance(e, HTTPException):
                return e
            raise e
        code = e.code if isinstance(e, HTTPException) else 500
        app.logger.exception(e)
        return jsonify({'error': 'Server error. Please try again.'}), code

    # Preload doc templates
    from utils.doc_engine import preload_templates
    with app.app_context():
        preload_templates()

    return app