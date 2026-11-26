<?php
// comum.php: o que todas as páginas PHP do site usam.
// A equipe muda só estas três linhas:
$NOME_SITE = 'Nome do negócio';
$COR = '#0b6e8a';
$SERVICOS = ['Pedir orçamento', 'Agendar horário', 'Tirar uma dúvida'];

// O banco de dados: um arquivo só, guardado FORA da pasta do site
$banco = new PDO('sqlite:/var/www/dados/site.db');
$banco->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
$banco->exec("CREATE TABLE IF NOT EXISTS cadastros (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  nome TEXT NOT NULL,
  telefone TEXT NOT NULL,
  servico TEXT NOT NULL,
  mensagem TEXT,
  criado_em TEXT DEFAULT (datetime('now', 'localtime'))
)");

// h(): mostra o texto do visitante como texto, nunca como código
function h($texto) {
  return htmlspecialchars((string) $texto, ENT_QUOTES, 'UTF-8');
}

// topo(): o começo de toda página PHP, com o menu
function topo($titulo) {
  global $NOME_SITE, $COR;
  echo '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">';
  echo '<meta name="viewport" content="width=device-width, initial-scale=1">';
  echo '<title>' . h($titulo) . ' - ' . h($NOME_SITE) . '</title>';
  echo '<style>
    body { font-family: Arial, sans-serif; margin: 0; background: #f4f4f4; color: #222; }
    header { background: ' . h($COR) . '; color: #fff; padding: 16px; }
    header strong { font-size: 22px; display: block; margin-bottom: 8px; }
    header a { color: #fff; margin-right: 16px; text-decoration: none; font-weight: bold; }
    main { max-width: 800px; margin: 24px auto; background: #fff; padding: 20px; border-radius: 10px; }
    label { display: block; margin-top: 14px; font-weight: bold; }
    input, select, textarea { width: 100%; padding: 10px; font-size: 16px; box-sizing: border-box; margin-top: 4px; }
    button, .botao { background: ' . h($COR) . '; color: #fff; border: 0; padding: 10px 18px; font-size: 16px;
      border-radius: 6px; margin-top: 14px; cursor: pointer; text-decoration: none; display: inline-block; }
    table { border-collapse: collapse; width: 100%; }
    td, th { border-bottom: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; }
    td form { display: inline; }
    .ok { background: #e6f6ea; padding: 12px; border-radius: 6px; }
    .erro { background: #fde8e8; padding: 12px; border-radius: 6px; }
  </style></head><body>';
  echo '<header><strong>' . h($NOME_SITE) . '</strong>';
  echo '<a href="index.html">Início</a><a href="servicos.html">Serviços</a>';
  echo '<a href="contato.html">Contato</a><a href="cadastro.php">Cadastre-se</a></header><main>';
}

// rodape(): o fim de toda página PHP
function rodape() {
  echo '</main></body></html>';
}
