sudo tar -czf ~/backup-servidor-equipe.tar.gz /var/www/html /var/www/dados /etc/nginx/sites-available/default
sudo chown $USER:$USER ~/backup-servidor-equipe.tar.gz
ls -lh ~/backup-servidor-equipe.tar.gz
