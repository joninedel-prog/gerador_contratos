tail -3 /var/www/html/comum.php
sudo tee -a /var/www/html/comum.php > /dev/null <<'FIM'

// ===== Encontro 3: login =====
session_start();

$banco->exec("CREATE TABLE IF NOT EXISTS usuarios (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  usuario TEXT NOT NULL UNIQUE,
  senha TEXT NOT NULL
)");

// exigir_login(): quem não entrou é mandado para o login
function exigir_login() {
  if (empty($_SESSION['usuario'])) {
    header('Location: login.php');
    exit;
  }
}
FIM
tail -20 /var/www/html/comum.php
