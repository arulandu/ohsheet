#!/bin/bash
echo "Setting up server..."
cd backend

if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt


echo "Setting up SSL certificates..."
brew install mkcert
mkcert -install
mkcert localhost 127.0.0.1 0.0.0.0::1

echo "Setting up plugin..."
cd ../plugin/ohsheet
defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
npm install

echo "Setup complete! 🎉"
echo ""
echo "To start development:"
echo "1. Backend: cd backend && source venv/bin/activate && python src/main.py"
echo "2. Frontend: cd plugin/ohsheet && npm start"
