sudo cp /etc/nginx/sites-available/default ~/nginx-default.original
sudo tee /etc/nginx/sites-available/default > /dev/null <<'FIM'
server {
    listen 80 default_server;
    listen [::]:80 default_server;
    root /var/www/html;
    index index.html index.php;
    server_name _;

    location / {
        try_files $uri $uri/ =404;
    }

    location ~ \.php$ {
        include snippets/fastcgi-php.conf;
        fastcgi_pass unix:/run/php/php-fpm.sock;
    }
}
FIM
sudo nginx -t
sudo systemctl reload nginx
