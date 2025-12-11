web: gunicorn supply_chain_analytics:server --timeout 120 --workers 2 --threads 4 --worker-class gthread --max-requests 500 --max-requests-jitter 50 --keep-alive 5 --log-level warning
