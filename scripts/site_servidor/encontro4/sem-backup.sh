sudo apt update
sudo apt install -y nginx
echo '<h1>Servidor VM da equipe no ar</h1>' | sudo tee /var/www/html/index.html
