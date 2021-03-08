# Introduction 
Versão para incluir a funcionalidade de gerar PDF individual com senha
Para isso houve mudança no formato do arquivo de índice, adicionando uma coluna nova e ficando desta forma

arqECM;pagInicial;pagFatura;pdfSenha
00000669760830492021030520210222;1;3;SENHA_DESEJADA_PARA_ESTE_ARQUIVO
00000669880015912021030520210226;4;3;NOPASSWORD
00000669840667622021031020210224;7;3;NOPASSWORD

Se a palavra NOPASSWORD estiver presente no campo de senha o PDF individual será gerado sem senha, caso contenha qualquer outra coisa o conteúdo será a senha aplicada.



