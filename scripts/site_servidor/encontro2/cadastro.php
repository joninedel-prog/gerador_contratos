<?php
require 'comum.php';

$aviso = '';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
  $nome = trim($_POST['nome'] ?? '');
  $telefone = trim($_POST['telefone'] ?? '');
  $servico = $_POST['servico'] ?? '';
  $mensagem = trim($_POST['mensagem'] ?? '');

  if ($nome === '' || $telefone === '' || !in_array($servico, $SERVICOS, true)) {
    $aviso = '<p class="erro">Preencha o nome, o telefone e escolha uma opção.</p>';
  } else {
    $sql = $banco->prepare('INSERT INTO cadastros (nome, telefone, servico, mensagem) VALUES (?, ?, ?, ?)');
    $sql->execute([$nome, $telefone, $servico, $mensagem]);
    $aviso = '<p class="ok">Recebemos o seu cadastro, ' . h($nome) . '! Logo entramos em contato.</p>';
  }
}

topo('Cadastre-se');
?>
<h1>Cadastre-se</h1>
<?= $aviso ?>
<form method="post">
  <label>Nome <input name="nome" required maxlength="80"></label>
  <label>Telefone <input name="telefone" required maxlength="20"></label>
  <label>O que você procura
    <select name="servico">
      <?php foreach ($SERVICOS as $s): ?>
        <option><?= h($s) ?></option>
      <?php endforeach; ?>
    </select>
  </label>
  <label>Mensagem <textarea name="mensagem" rows="4" maxlength="500"></textarea></label>
  <button type="submit">Enviar</button>
</form>
<?php rodape(); ?>
