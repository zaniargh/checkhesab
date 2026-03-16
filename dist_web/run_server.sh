#!/bin/bash
echo "Installing requirements..."
pip install -r requirements.txt

echo "Starting server on port 80..."
# Using sudo to bind to port 80 if necessary, or running on 8000
uvicorn app:app --host 0.0.0.0 --port 8000
