![:hiper logo](https://cdn.vikomet.eu/hiperstream/filestorage/7CDC6DADC45843E4898D3844241F3001.png)
<br>
[![Build Status](https://img.shields.io/badge/build-sucess-blue)](https://github.com/hiperstream/hiper-break-pdf-files)
[![Open in Visual Studio Code](https://open.vscode.dev/badges/open-in-vscode.svg)](https://open.vscode.dev/hiperstream/hiper-break-pdf-files)

# hiper-break-pdf-files

Código que quebra um pdf já existente e proteje com senha o arquivo gerado.
Para isso houve mudança no formato do arquivo de índice, adicionando uma coluna nova e ficando desta forma

## Pré Requisitos
* [dotnet 4.6](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net462)
* [VSCode](https://code.visualstudio.com/download)
* [iText](https://itextpdf.com/)

## Exemplo

```
arqECM;pagInicial;pagFatura;pdfSenha
00000669760830492021030520210222;1;3;SENHA_DESEJADA_PARA_ESTE_ARQUIVO
00000669880015912021030520210226;4;3;NOPASSWORD
00000669840667622021031020210224;7;3;NOPASSWORD
```

Quando a palavra SENHA é usada no campo JOB da aplicação é obrigatória a existência da coluna pdfSenha com as senhas desejadas para cada PDF individual.

Se não houver a palavra SENHA a coluna pdfSenha é opcional e será ignorada caso exista.




