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
    </ul>
  <h3>Exemplos de Restrições:</h3>
    <p><strong><em>Arquivos suportados pelo programa:</em></strong></p>
    <ul>
      <li><em> .xlsx</em></li>
      <li><em>.xlsm</em></li>
      <li><em>.xltx</em></li>
      <li><em>.xltm</em></li>
    </ul>
    <p><strong><em>Formatação de tags suportada pelo programa:</em></strong></p>
       <ul>
      <li><em> PB-6027SA-02</em></li>
      <li><em> PB-6027SA-002</em></li>
      <li><em> PB-6027SA-0150</em></li>
      <li><em> "-" para equipamentos não taggeados</em></li>
    </ul>
    <p><strong><em>Obs: Sempre antes do primeiro "-" devem haver de 2 a 3 letras, de 5 a 10 algarismos alfanúmericos para identificação</br>
      da área e depois do segundo "-" sempre deve haver de 2 a 4 números, sendo que só pode haver no máximo 2 zeros ex : PB-6027SA-0002,</br>
      a formtação acima é inválida, entretanto, "PB-6027SA-002", "PB-6027SA-02" ou "PB-6027SA-2" são formatações válidas. Além disso,</br>
       vale a pena ser dito que a formatação descrita acima é válida tanto para a PQ quanto para LE, todavia, na PQ não necessidade de </br>
        inserir o identificador de área, o próprio código se encarregará disso. Exemplos de formatação correta para PQ são: "PB-02", </br>
        "PB-002" e "PB-2".</em></strong></p>
        
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
           a extensão correta é (.xlsx ou .xlsm). Caso algum dos arquivos não esteja na extensão correta, abra ele, vá em salvar</br>
           como e salve com a extensão correta. Por úlitmo, tendo certeza que as extensões e os caminhos estão</br>
           corretos clique em "Submit" e espere.</p></em>
    <h2>Considerações Finais</h2>
      <p><em> Portanto, com todas as devidas explicações feitas, é relevante ser frisado que toda a lógica foi desenvolvida para o modelo Vale,</b>
      para funcionar em outros modelos,<strong> talvez</strong> seja necesário alterações no código.</em></p>
           
          
