# Introduction 
Versão para incluir a funcionalidade de gerar PDF individual com senha
Para isso houve mudança no formato do arquivo de índice, adicionando uma coluna nova e ficando desta forma

arqECM;pagInicial;pagFatura;pdfSenha
00000669760830492021030520210222;1;3;SENHA_DESEJADA_PARA_ESTE_ARQUIVO
00000669880015912021030520210226;4;3;NOPASSWORD
00000669840667622021031020210224;7;3;NOPASSWORD

Quando a palavra SENHA é usada no campo JOB da aplicação é obrigatória a existência da coluna pdfSenha com as senhas desejadas para cada PDF individual.

Se não houver a palavra SENHA a coluna pdfSenha é opcional e será ignorada caso exista.




