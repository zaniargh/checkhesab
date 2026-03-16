#!/bin/bash
echo "Installing Python dependencies and Nginx..."
sudo apt update
sudo apt install -y python3-pip python3-venv nginx

echo "Setting up application directory..."
sudo mkdir -p /var/www/hesab
sudo cp -r * /var/www/hesab/
cd /var/www/hesab

echo "Installing Python Requirements globally..."
sudo pip3 install -r requirements.txt --break-system-packages

echo "Configuring Nginx..."
sudo cp hesab.gitzan.com.conf /etc/nginx/sites-available/
sudo ln -sf /etc/nginx/sites-available/hesab.gitzan.com.conf /etc/nginx/sites-enabled/
sudo rm -f /etc/nginx/sites-enabled/default
sudo systemctl restart nginx

echo "Setting up SystemD Service..."
sudo cp hesab.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable hesab
sudo systemctl restart hesab

echo "=========================================================="
echo "Deployment Complete! The app is now running on your server."
echo "You can access it at: http://hesab.gitzan.com"
echo "=========================================================="
