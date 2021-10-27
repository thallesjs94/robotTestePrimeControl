Requisitos
==========
O projeto utiliza alguma Libs do Python que são pré-requisitos para executar, sendo elas:

os, selenium, pathlib, re, time, openpyxl, PyPDF2 e csv

Utilizei a versão 93.0 do Firefox para o desenvolvimento. Caso execute em alguma outra versão, pode haver problema de compatibilidade do Geckodriver.

Para instalar
=============
`pip install nomedalib`

Para executar
=============
OBS: Dentro da pasta do processo
`python main.py`

Decisões tomadas
================

- Não estava claro no processo o que fazer com as informações extraídas do PDF, por isso só realizei a extração, porém não fiz a comparação.
- Optei por desenvolver a automação utilizando o Python puro, sem utilizar nenhum framework