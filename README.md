# Leitura_XML
Leitura das principais informações do XML de documentos fiscais

Utilizei as bibliotecas os, tkinter, pandas, xmltodict e tqdm para criar um código que pede para selecionar uma pasta que contenha arquivos com a extensão ".xml (utilizei uma janela do "tkinter" e "os" para listar os arquivos) e faça a leitura das informações (utilizei o "xmltodict" para transformar o XML em um dicionário) que contem naquela nota fiscal.

Encontrei algumas diferenças na estrutura dos arquivos XML das notas fiscais e precisei separar o código em 4 funções, as duas primeiras para leitura de NFE (Nota Fiscal Eletrônica), a terceira faz a leitura das informações dos CTe (Conhecimento de Transporte Eletrônico) e a quarta para leitura de NFCE (Nota Fiscal de Consumidor Eletrônica):
A primeira função lê arquivos XML de NFE solicitados na Receita do Estado do Paraná.
A segunda função lê os arquivos XML emitidos pelas empresas ou que foram baixados dos sites Portal da NFE (Site do Governo Federal necessita de certificado digital) ou Fsist (algumas notas não necessitam de certificado digital para baixar o XML).
A terceira função faz a leitura dos arquivos XML de notas de transporte (CTe).
A quarta função faz a leitura dos arquivos XML de notas fiscais emitidas para consumidor final.

As funções fazem basicamente a mesma leitura do XML, respeitando as diferenças que os XML's possuem em sua estrutura.
As funções irão retornar um dicionário com as informações do XML da nota fiscal lida, a partir desses dicionários irá criar um DataFrame da biblioteca pandas e a partir do DF do pandas irá gerar um arquivo Excel com as informações mais relevantes lidas do XML.

Obs: Caso não consiga ler o XML altere as vezes em que aparecer "documento = xmltodict.parse(arquivo)" para "documento = xmltodict.parse(arquivo, encoding= 'ANSI')", no mês 08/2022 houve alguma alteração nos documentos solicitados na Receita PR e alguns XML vieram com a codificação ANSI.
