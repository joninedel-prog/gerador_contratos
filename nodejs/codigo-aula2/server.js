// Módulos que já vêm com o Node.js
const http = require('http');
const fs = require('fs');
const path = require('path');

// Porta onde o servidor vai atender
const PORTA = 3000;

// O "garçom": roda a cada pedido que chega
const servidor = http.createServer((pedido, resposta) => {
  const hora = new Date().toLocaleTimeString('pt-BR');
  console.log(`${hora} - ${pedido.socket.remoteAddress} pediu ${pedido.url}`);

  const arquivo = path.join(__dirname, 'index.html');
  fs.readFile(arquivo, (erro, conteudo) => {
    if (erro) {
      resposta.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      resposta.end('Erro: não consegui ler o index.html');
      return;
    }
    resposta.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    resposta.end(conteudo);
  });
});

// Liga o servidor e espera pedidos de qualquer placa de rede
servidor.listen(PORTA, '0.0.0.0', () => {
  console.log(`Servidor rodando na porta ${PORTA}`);
});
