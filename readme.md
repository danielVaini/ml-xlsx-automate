# Automação planilha Mercado livre
  Este projeto tem como objetivo automatizar a planilha de anúncios em massa fornecida pelo mercado livre.

  (Anúmcios em massa do mesmo produto com modelos diferentes)

## Como instalar
  Descompactar o arquivo `planilha-mp.rar` em um local de sua preferência e executar o arquivo `planilha-mp.exe` para executar o software.
  
  ![Captura de tela 2023-12-27 074205](https://github.com/danielVaini/ml-xlsx-automate/assets/57446226/63bb131c-06f0-4305-9eb6-9a8c14ba4e06)

## Como utilizar
  O primeiro passo, é entrar em sua conta do [**Mercado Livre**](https://www.mercadolivre.com.br/) e ir na opção de **Anunciador em massa** onde o mercado livre irá disponibilizar a planilha para download. Já com a planilha, basta carrega-lá para o programa com o botão `Abrir Planilha ML`.

  Para anunciar um mesmo produto que possui diversos modelos diferentes, precisará fornecer ao software uma segunda planilha onde contem todos os modelos dísponiveis para este produto, para isso basta utilizar o botão `Abrir Planilha MODELOS`.

  **EXEMPLO DE PALNILHA COM MODELOS**
  
  ![Captura de tela 2023-12-17 163802](https://github.com/danielVaini/ml-xlsx-automate/assets/57446226/5632747f-8301-4b3d-b9d3-c401afeff567)

### Título do anúncio
  O título deve conter no local do modelo da peça um “_”  para que o programa altere todos os modelos do título.

  Exemplo de título: 
  - Como é: Celular nokia modelo 123 tijolo
  - Como fica: Farol gol modelo _ tijolo

### Link fotos
  Tem um padrão a ser seguido, no caso, seguir a documentação do Mercado Livre e adicionar suas fotos em um link, colar esse link no campo `Link fotos`.

### Campos especiais
  O campo `Iniciar na linha` caso informado, será adicionado os produtos a partir da linha desejada, caso queira utilizar a mesma planilha para diversos tipos de anúncio. No caso, precisa abrir a planilha criada e verificar em qual linha deve ser iniciada a contagem novamente. 
  
  CASO VÁ UTILIZAR A MESMA PLANILHA PARA DIVERSOS ANÚNCIOS DIFERENTES, LEMBRE SE DE CARREGAR A PLANILHA NOVA CRIADA NA PRIMEIRA EXECUÇÃO ONDE JÁ CONTÉM ANÚNCIOS CRIADOS.
  
  
  `Preço max`, caso informado, será utilizado o campo `Preço` como **preço mínimo** e o campo `Preço max` como **preço máximo** para criar variações de preços entre os dois valores informados, sendo eles inclusos e colocados na planilha.
  
  `Variar frete`, caso deseje alterar entre frete grátis ou por conta do comprador.

### Adicionar os dados

  Para adicionar os dados, após todos os passos seguidos, basta pressionar  o botão `Adicionar dados` e em seguida `Salvar planilha`.

Salve a planilha no local de sua preferência.
