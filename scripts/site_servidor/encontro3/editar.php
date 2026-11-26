<?php
require 'comum.php';
exigir_login();

$id = (int) ($_GET['id'] ?? 0);
$sql = $banco->prepare('SELECT * FROM cadastros WHERE id = ?');
$sql->execute([$id]);
$c = $sql->fetch(PDO::FETCH_ASSOC);
if (!$c) {
  header('Location: painel.php');
  exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
  $sql = $banco->prepare('UPDATE cadastros SET nome = ?, telefone = ?, servico = ?, mensagem = ? WHERE id = ?');
  $sql->execute([
    trim($_POST['nome'] ?? ''),
    trim($_POST['telefone'] ?? ''),
    $_POST['servico'] ?? '',
    trim($_POST['mensagem'] ?? ''),
    $id,
  ]);
  header('Location: painel.php');
  exit;
}

topo('Editar');
?>
<h1>Editar o cadastro #<?= $c['id'] ?></h1>
<form method="post">
  <label>Nome <input name="nome" required maxlength="80" value="<?= h($c['nome']) ?>"></label>
  <label>Telefone <input name="telefone" required maxlength="20" value="<?= h($c['telefone']) ?>"></label>
  <label>Opção
    <select name="servico">
      <?php foreach ($SERVICOS as $s): ?>
        <option <?= $s === $c['servico'] ? 'selected' : '' ?>><?= h($s) ?></option>
      <?php endforeach; ?>
    </select>
  </label>
  <label>Mensagem <textarea name="mensagem" rows="4" maxlength="500"><?= h($c['mensagem']) ?></textarea></label>
  <button type="submit">Salvar</button>
  <a class="botao" href="painel.php">Voltar</a>
</form>
<?php rodape(); ?>
