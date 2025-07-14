# xlwings addin install
# npm-install -g yo generator
# yo office
brew install mkcert
mkcert -install
mkcert localhost 127.0.0.1 0.0.0.0::1
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
cp ./plugin/ohsheet/manifest.json ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
# rm -rf ~/Library/Containers/com.microsoft.Excel/Data
# defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
