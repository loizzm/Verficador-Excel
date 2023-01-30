<h1>Documentação</h1>
<h2>Introdução</h2>
<p><em>Em diversas ocasiões temos que comparar listas de equipamentos com planilha de quantidades, além de ser um tarefa demorada,</br>
requer muita atenção. Nesse contexto, esse recurso funciona como um verificador de equidade entre as duas planilhas. Tendo como objetivo</br>
uma economia de tempo e uma verificação mais consistente</em></p>
<h2>Restrições</h2>
<p><em>Bom, para compreender o funcionamento do código é importante ser dito o que <strong>não</strong> dever ser feito com ele.</em></p>
<p><em>Restrições da aplicação</em></p>
<ul>
  <li><em>De forma alguma os arquivos usados para comparação devem estar no formato <strong>.xls</strong></em></li>
  <li><em>O caminho de um arquivo e de outro deve sempre estar separado por uma quebra de linha</em></li>
  <li><em>O código descrito <strong>ainda</strong> não possui a função de verificar equipamentos não taggeados</em></li>
</ul>
<h3>Exemplos de Restrições:</h3>
<p><strong><em>Formatação do arquivos Files.txt:</em></strong></p>
<img src="https://user-images.githubusercontent.com/114931499/215488739-5f6c0790-4a95-464b-bf59-1f4346c05ba8.png" width="200px" />
<p><strong><em>Arquivos suportados pelo programa:</em></strong></p>
<ul>
  <li><em> .xlsx</em></li>
  <li><em>.xlsm</em></li>
  <li><em>.xltx</em></li>
  <li><em>.xltm</em></li>
</ul>
<h2>Passo a Passo</h2>
<p><em>Primeiramente, é importante ser dito que para um melhor encapsulamneto do código e para uma diminuição na quantidade de possíveis</br>
erros, não haverá interação direta com o código. A aplicação funcionará da seguinte forma, será baixado um executável(.exe)</br>
 e um arquivo chamadoFiles.txt(onde será colocado o caminho dos arquivos). Feito isso, o usário só terá que executar o arquivo e um pop-up aparecerá alertando-o de inconsistências encontradas entre os arquivos.</em></p>
