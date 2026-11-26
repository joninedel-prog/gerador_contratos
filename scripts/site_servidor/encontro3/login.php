<?php
require 'comum.php';

$primeiro_acesso = ($banco->query('SELECT COUNT(*) FROM usuarios')->fetchColumn() == 0);
$erro = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
  $usuario = trim($_POST['usuario'] ?? '');
  $senha = $_POST['senha'] ?? '';

  if ($primeiro_acesso) {
    if ($usuario === '' || strlen($senha) < 8) {
      $erro = 'Escolha um usuário e uma senha com 8 caracteres ou mais.';
    } else {
      $sql = $banco->prepare('INSERT INTO usuarios (usuario, senha) VALUES (?, ?)');
      $sql->execute([$usuario, password_hash($senha, PASSWORD_DEFAULT)]);
      session_regenerate_id(true);
      $_SESSION['usuario'] = $usuario;
      header('Location: painel.php');
      exit;
    }
  } else {
    $sql = $banco->prepare('SELECT senha FROM usuarios WHERE usuario = ?');
    $sql->execute([$usuario]);
    $hash = $sql->fetchColumn();
    if ($hash && password_verify($senha, $hash)) {
      session_regenerate_id(true);
      $_SESSION['usuario'] = $usuario;
      header('Location: painel.php');
      exit;
    }
    $erro = 'Usuário ou senha errados.';
  }
}

topo('Entrar');
?>
<h1><?= $primeiro_acesso ? 'Primeiro acesso: crie o administrador' : 'Entrar no painel' ?></h1>
<?php if ($erro): ?><p class="erro"><?= h($erro) ?></p><?php endif; ?>
<form method="post">
  <label>Usuário <input name="usuario" required maxlength="30"></label>
  <label>Senha <input name="senha" type="password" required></label>
  <button type="submit"><?= $primeiro_acesso ? 'Criar e entrar' : 'Entrar' ?></button>
</form>
<?php rodape(); ?>
