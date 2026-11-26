<?php
require 'comum.php';
exigir_login();

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['apagar'])) {
  $sql = $banco->prepare('DELETE FROM cadastros WHERE id = ?');
  $sql->execute([(int) $_POST['apagar']]);
  header('Location: painel.php');
  exit;
}

$busca = trim($_GET['busca'] ?? '');
$sql = $banco->prepare('SELECT * FROM cadastros WHERE nome LIKE ? ORDER BY id DESC');
$sql->execute(['%' . $busca . '%']);
$lista = $sql->fetchAll(PDO::FETCH_ASSOC);

topo('Painel');
?>
<h1>Painel de cadastros</h1>
<p>Você entrou como <b><?= h($_SESSION['usuario']) ?></b> · <a href="sair.php">Sair</a></p>
<form method="get">
  <label>Procurar pelo nome <input name="busca" value="<?= h($busca) ?>"></label>
  <button type="submit">Procurar</button>
</form>
<p><?= count($lista) ?> cadastro(s).</p>
<table>
  <tr><th>#</th><th>Quando</th><th>Nome</th><th>Telefone</th><th>Opção</th><th>Mensagem</th><th></th></tr>
  <?php foreach ($lista as $c): ?>
  <tr>
    <td><?= $c['id'] ?></td>
    <td><?= h($c['criado_em']) ?></td>
    <td><?= h($c['nome']) ?></td>
    <td><?= h($c['telefone']) ?></td>
    <td><?= h($c['servico']) ?></td>
    <td><?= h($c['mensagem']) ?></td>
    <td>
      <a class="botao" href="editar.php?id=<?= $c['id'] ?>">Editar</a>
      <form method="post" onsubmit="return confirm('Apagar este cadastro?')">
        <button name="apagar" value="<?= $c['id'] ?>">Apagar</button>
      </form>
    </td>
  </tr>
  <?php endforeach; ?>
</table>
<?php rodape(); ?>
