<?php
require 'comum.php';

$lista = $banco->query('SELECT * FROM cadastros ORDER BY id DESC')->fetchAll(PDO::FETCH_ASSOC);

topo('Painel');
?>
<h1>Painel de cadastros</h1>
<p><?= count($lista) ?> cadastro(s) no banco.</p>
<table>
  <tr><th>#</th><th>Quando</th><th>Nome</th><th>Telefone</th><th>Opção</th><th>Mensagem</th></tr>
  <?php foreach ($lista as $c): ?>
  <tr>
    <td><?= $c['id'] ?></td>
    <td><?= h($c['criado_em']) ?></td>
    <td><?= h($c['nome']) ?></td>
    <td><?= h($c['telefone']) ?></td>
    <td><?= h($c['servico']) ?></td>
    <td><?= h($c['mensagem']) ?></td>
  </tr>
  <?php endforeach; ?>
</table>
<?php rodape(); ?>
