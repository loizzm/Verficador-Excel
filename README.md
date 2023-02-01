<h1>Documentação</h1>
  <h2>Introdução</h2>
    <p><em>Em diversas ocasiões temos que comparar listas de equipamentos com planilha de quantidades, além de ser um tarefa</br>
    demorada, requer muita atenção. Nesse contexto, esse recurso funciona como um verificador de equidade entre as duas</br>
    planilhas. Tendo como objetivo uma economia de tempo e uma verificação mais consistente</em></p>
  <h2>Restrições</h2>
    <p><em>Bom, para compreender o funcionamento do código é importante ser dito o que <strong>não</strong> dever ser feito com ele.</em></p>
    <p><em>Restrições da aplicação</em></p>
    <ul>
      <li><em>De forma alguma os arquivos usados para comparação devem estar no formato <strong>.xls</strong></em></li>
      <li><em>O código descrito <strong>ainda</strong> não possui a função de verificar equipamentos não taggeados</em></li>
    </ul>
  <h3>Exemplos de Restrições:</h3>
    <p><strong><em>Arquivos suportados pelo programa:</em></strong></p>
    <ul>
      <li><em> .xlsx</em></li>
      <li><em>.xlsm</em></li>
      <li><em>.xltx</em></li>
      <li><em>.xltm</em></li>
    </ul>
  <h2>Passo a Passo</h2>
    <p><em>Primeiramente, é importante ser dito que para um melhor encapsulamneto do código e para uma diminuição na quantidade</br>
    de possíveis erros, não haverá interação direta com o código. A aplicação funcionará da seguinte forma, será baixado</br>
     um executável(.exe) e nele será inserido os caminhos para os arquivos desejados e o código se encarregará do resto. Feito isso,</br>
     o usário só terá que executar o  arquivo e um pop-up    aparecerá alertando-o de inconsistências encontradas entre os arquivos.</em></p>
