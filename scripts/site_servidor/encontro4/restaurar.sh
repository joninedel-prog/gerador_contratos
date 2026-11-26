sudo apt update
sudo apt install -y nginx php-fpm php-cli php-sqlite3 sqlite3
cd /
sudo tar -xzvf ~/backup-servidor-equipe.tar.gz
sudo chown -R www-data:www-data /var/www/dados
sudo nginx -t
sudo systemctl reload nginx
cd ~
