"""
Gunicorn configuration for single-process deployment
Optimized for 20 internal users on office network
"""

import multiprocessing

# Single worker process to avoid multi-process state synchronization issues
workers = 1

# Thread pool for concurrent request handling
threads = 4

# Use threaded worker class
worker_class = "gthread"

# Timeout for requests (2 minutes)
timeout = 120

# Binding address
bind = "0.0.0.0:8000"

# Logging
loglevel = "info"
accesslog = "-"
errorlog = "-"

# Graceful timeout for worker shutdown
graceful_timeout = 30

# Keep-alive
keepalive = 5

# Worker restart configuration
max_requests = 1000
max_requests_jitter = 50
