# WSH_BrowserInterface

Portugês:  

Quando um Script WSH (Windows Script Host)está em execução não sabemos como está o processamento, ou não conseguimos pausa-lo ou finaliza-lo de maneira elegante, então criamos uma interface com o IE (Internet Explorer) para interagir com o Script WSH.

Primeiramente o Script WSH abre uma janela do Navegador IE e cria uma página HTML;
Em seguida o Script lê periodicamente na página a interação do usuário;

Conforme a interação, o Script Pausa ou Continua ou Finaliza o seu processamento.
O sincronismo entre o Script e a Página HTML é obtida por meio do próprio Script através da leitura periódica 
do conteúdo de uma TAG HTML oculta na página. Por sua vez, essa TAG é modificada conforme a interação do usuário com a página HTML. 

A Interface do IE é composta por:
- 3 botões: Continuar, Pausar e Finalizar;  
- 2 Textos alterados pelo Script periodicamente;

Observação 1: O seu Script WSH principal precisa estar adaptado para periodicamente realizar a leitura da página HTML. Isso pode ser obtido inserindo uma verificação em pontos não cruciais do seu processamento.


English: 

When a WSH Script (Windows Script Host) is running we do not know how it is processing, or we were unable to pause or finish it in an elegant way, so we created an interface with IE (Internet Explorer) to interact with the WSH Script .   

First, the WSH Script opens an IE Browser window and creates an HTML page;
Then the Script periodically reads the user's interaction on the page;

Depending on the interaction, the Script Pauses or Continues or Finalizes its processing.
The synchronism between the Script and the HTML page is obtained through the Script itself through periodic reading
of the content of an HTML tag hidden on the page. In turn, this TAG is modified according to the user's interaction with the HTML page.

The IE Interface is composed of:
- 3 buttons: Continue, Pause and Finish;
- 2 Texts changed by the Script periodically;

Note 1: Your main WSH Script needs to be adapted to periodically read the HTML page. This can be achieved by inserting a check at non-crucial points in its processing.

