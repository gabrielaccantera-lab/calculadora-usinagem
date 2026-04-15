# Como publicar o app — passo a passo (sem conhecimento técnico)

Você vai precisar só de um navegador. Nenhuma instalação necessária.

---

## Passo 1 — Criar conta no GitHub

1. Acesse https://github.com
2. Clique em "Sign up"
3. Crie a conta com seu e-mail (pode ser qualquer e-mail)
4. Confirme o e-mail quando chegar na sua caixa de entrada

---

## Passo 2 — Criar o repositório e subir os arquivos

1. Após entrar no GitHub, clique no "+" no canto superior direito
2. Clique em "New repository"
3. Em "Repository name" escreva: `calculadora-usinagem`
4. Deixe marcado como "Public"
5. Marque a caixinha "Add a README file"
6. Clique em "Create repository"

Agora suba os arquivos:

7. Na página do repositório, clique em "Add file" → "Upload files"
8. Arraste os 3 arquivos de uma vez: `app.py`, `requirements.txt`, `README.md`
9. Lá embaixo clique em "Commit changes"

---

## Passo 3 — Criar conta no Streamlit Cloud

1. Acesse https://share.streamlit.io
2. Clique em "Sign up"
3. Clique em "Continue with GitHub" — usa a mesma conta que você acabou de criar
4. Autorize o acesso quando pedir

---

## Passo 4 — Publicar o app

1. No Streamlit Cloud clique em "New app"
2. Em "Repository" selecione `calculadora-usinagem`
3. Em "Branch" deixe `main`
4. Em "Main file path" escreva `app.py`
5. Clique em "Deploy!"
6. Aguarde cerca de 2 minutos

Pronto — vai aparecer um link do tipo:
**https://calculadora-usinagem.streamlit.app**

Esse link é público. Qualquer pessoa com o link consegue acessar e usar o app, sem precisar instalar nada.

---

## Como usar o app

1. Abra o link no navegador
2. Clique em "Browse files" e selecione o arquivo `.xlsm`
3. O app lê tudo automaticamente e mostra os resultados

---

## Dúvidas frequentes

**O link some depois de um tempo?**
Não — fica no ar enquanto a conta existir. Se ficar sem acesso por 7 dias o Streamlit "hiberna" o app, mas basta abrir o link que ele volta em segundos.

**Posso atualizar o app depois?**
Sim — basta substituir o arquivo `app.py` no GitHub (mesmo caminho: "Add file" → "Upload files") e o Streamlit atualiza sozinho em minutos.

**Outras pessoas conseguem ver meus dados?**
Não — cada usuário que acessa faz o upload do próprio arquivo. Nenhum dado fica salvo no servidor.
