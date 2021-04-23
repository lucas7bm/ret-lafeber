# Ajustes e Apuração RET - Lafeber
Este script faz o cálculo da sub-apuração de ICMS do Regime Especial de Tributação (RET) da empresa Lafeber.
Além disso, faz os lançamentos dos registros de ajuste de estornos no registro C197 da EFD e cria os registros de sub-apuração 1900 e seus filhos, seguindo a especificação do documento legal que permitiu a adoção do regime especial.

#Como gerar um executável
Uma GUI foi criada usando PySimpleGUI e o o código todo pode ser transformado em um executável usando a biblioteca pyinstaller.
Para gerar o executável, basta executar o seguinte comando:

```pyinstaller .\Ajustes RET Lafeber.py --onefile --noupx --noconsole -i .\icon.ico```

#Dependências
 - PySimpleGUI
 - XLSXWriter
