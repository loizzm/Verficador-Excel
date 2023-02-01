<h1>Documentação</h1>
  <h2>Introdução</h2>
    <p><em>Em diversas ocasiões temos que comparar listas de equipamentos com planilha de quantidades, além de ser um tarefa</br>
    demorada, requer muita atenção. Nesse contexto, esse recurso funciona como um verificador de equidade entre as duas</br>
    planilhas. Tendo como objetivo uma economia de tempo e uma verificação mais consistente.</em></p>
  <h2>Restrições</h2>
    <p><em>Bom, para compreender o funcionamento do código é importante ser dito o que <strong>não</strong> dever ser feito com ele.</em></p>
    <p><em>Restrições da aplicação:</em></p>
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
    <p><em>Primeiramente, é importante ser dito que para um melhor encapsulamneto do código e para uma diminuição na </br>
    quantidade de possíveis erros, não haverá interação direta com o código. A aplicação funcionará da seguinte forma, será baixado</br>
     um executável(.exe) e nele será inserido os caminhos para os arquivos desejados e o código se encarregará do resto. Feito isso,</br>
     o usário só terá que executar o  arquivo e um pop-up    aparecerá alertando-o de inconsistências encontradas</br>
     entre os arquivos.</em></p>
      <h3>Demonstração:</h3>
          <p><em>Primeiro, clique duas vezes no arquivo <strong>main.exe</strong>. Feito isso, você verá essa tela:</p></em>
          <img src="https://user-images.githubusercontent.com/114931499/216100902-f3c05b1a-890d-4e6e-92c9-53313541881f.png" width"400px">
          <p><em>Agora, insira os caminhos dos arquivos em seus respectivos lugares, mas antes lembre-se que</br>
           além do caminho do arquivo, você deve adicionar uma "\" e o nome do arquivo desejado com a extensão</br>
           correta(.xlsx ou .xlsm). Caso algum dos arquivos não esteja na extensão correta, abra ele, vá em salvar</br>
           como e salve com a extensão correta. Por úlitmo, tendo certeza que as extensões e os caminhos estão</br>
           corretos clique em "Submit" e espere.</p></em>
