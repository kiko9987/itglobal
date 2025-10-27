#!/usr/bin/env python
"""Server startup script using Flask application factory pattern"""
import sys
import os

# Add dashboard directory to Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
dashboard_dir = os.path.join(current_dir, 'dashboard')
sys.path.insert(0, dashboard_dir)

if __name__ == '__main__':
    # Use factory pattern to create app with SocketIO
    from dashboard import create_app

    # Create app instance with development configuration for hot reload
    app, socketio = create_app('development')

    # Get port from environment or use default
    port = int(os.environ.get('PORT', 5000))
    debug = True  # Enable hot reload

    # Use SocketIO if available, otherwise fallback to regular Flask
    # Hot reload enabled - safe now that global cache variables are removed
    # 전역 변수 캐시 제거 완료로 핫리로드 안전하게 활성화 가능
    if socketio:
        print(f"Starting server with SocketIO on port {port} (Hot Reload Enabled)")
        socketio.run(app, host='0.0.0.0', port=port, debug=debug, use_reloader=True, allow_unsafe_werkzeug=True)
    else:
        print(f"Starting server without SocketIO on port {port} (Hot Reload Enabled)")
        app.run(host='0.0.0.0', port=port, debug=debug, use_reloader=True)