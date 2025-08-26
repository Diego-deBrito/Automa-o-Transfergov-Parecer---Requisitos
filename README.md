<h1>Robô de Análise de Propostas</h1>

<h2>1. Sobre o Projeto</h2>
<p>
Este projeto é um <strong>robô de automação em Python</strong> desenvolvido para analisar propostas dentro de um sistema web.  
O script automatiza a navegação, coleta dados relevantes (pareceres, requisitos e documentos) e gera uma planilha consolidada com a situação de cada proposta.  
Além disso, aplica <strong>regras de negócio</strong> para indicar a ação necessária em cada caso, reduzindo tempo e eliminando verificações manuais repetitivas.
</p>

<hr>

<h2>2. Funcionalidades</h2>
<ul>
  <li>Automação Web com <strong>Selenium</strong> para acessar e extrair informações.</li>
  <li>Tratamento e organização de dados com <strong>Pandas</strong>.</li>
  <li>Exportação dos resultados em formato <strong>Excel</strong>.</li>
  <li>Aplicação de regras de negócio para classificação automática:
    <ul>
      <li>Parceria Celebrada</li>
      <li>Em Celebração</li>
      <li>Proposta Rejeitada</li>
      <li>Entidade Pendência de Documentação</li>
      <li>Técnico Analisar</li>
    </ul>
  </li>
  <li>Mecanismo de <strong>checkpoint</strong> e <strong>reprocessamento de falhas</strong> para maior robustez.</li>
</ul>

<hr>

<h2>3. Estrutura do Projeto</h2>
<pre>
projeto-automacao-propostas/
│── src/
│   ├── Parecer-Requisitos.py   # Script principal
│── input/
│   ├── propostas.xlsx          # Planilha de entrada (exemplo fictício)
│── output/
│   ├── resultado.xlsx          # Saída gerada pelo robô
│── requirements.txt            # Dependências do projeto
│── README.md                   # Documentação
</pre>

<hr>

<h2>4. Tecnologias Utilizadas</h2>
<ul>
  <li>Python 3.10+</li>
  <li>Selenium</li>
  <li>BeautifulSoup</li>
  <li>Pandas</li>
  <li>OpenPyXL</li>
  <li>Webdriver Manager</li>
</ul>

<hr>

<h2>5. Instalação e Execução</h2>

<h3>Passo 1 – Clonar o repositório</h3>
<pre><code class="language-bash">
git clone https://github.com/seu-usuario/projeto-automacao-propostas.git
cd projeto-automacao-propostas
</code></pre>

<h3>Passo 2 – Criar ambiente virtual (opcional)</h3>
<pre><code class="language-bash">
python -m venv venv
</code></pre>

<h3>Passo 3 – Ativar ambiente virtual</h3>
<p>Windows:</p>
<pre><code class="language-bash">
venv\Scripts\activate
</code></pre>
<p>Linux/Mac:</p>
<pre><code class="language-bash">
source venv/bin/activate
</code></pre>

<h3>Passo 4 – Instalar dependências</h3>
<pre><code class="language-bash">
pip install -r requirements.txt
</code></pre>

<h3>Passo 5 – Iniciar Chrome em modo de depuração</h3>
<pre><code class="language-bash">
chrome.exe --remote-debugging-port=9222 --user-data-dir="C:/chrome_dev"
</code></pre>

<h3>Passo 6 – Executar o script</h3>
<pre><code class="language-bash">
python src/Parecer-Requisitos.py
</code></pre>

<h3>Menu de Execução</h3>
<pre>
[1] Processamento completo
[2] Reprocessar falhas
[3] Sair
</pre>

<hr>

<h2>6. Exemplo de Uso</h2>

<h3>Entrada (planilha fictícia)</h3>
<table>
  <tr><th>Nº Proposta</th><th>Instrumento</th><th>Situacional</th><th>Técnico Responsável</th></tr>
  <tr><td>12345</td><td>Convênio</td><td>Em Celebração</td><td>João Silva</td></tr>
</table>

<h3>Saída (gerada pelo robô)</h3>
<table>
  <tr><th>Nº Proposta</th><th>Instrumento</th><th>Ação Necessária</th><th>Certidões</th><th>Declarações</th><th>Histórico (Evento)</th></tr>
  <tr><td>12345</td><td>Convênio</td><td>Parceria Celebrada</td><td>OK</td><td>OK</td><td>Análise Registrada</td></tr>
</table>

<hr>

<h2>7. Aprendizados Demonstrados</h2>
<ul>
  <li>Automação de processos com Python (RPA).</li>
  <li>Extração, transformação e organização de dados.</li>
  <li>Implementação de regras de negócio em sistemas automatizados.</li>
  <li>Geração de relatórios estruturados em Excel.</li>
  <li>Estruturação de projetos e boas práticas de desenvolvimento.</li>
</ul>

<hr>

<h2>8. Observações</h2>
<p>
Os dados utilizados nos exemplos deste repositório são fictícios.  
O projeto foi adaptado para fins de demonstração e não expõe informações reais.
</p>
