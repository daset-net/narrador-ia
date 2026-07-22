# NarradorAI 🎙️

Aplicação web que analisa cada slide de uma apresentação PowerPoint visualmente e gera notas de narrador enriquecidas em português do Brasil usando a API do Groq.

## Como funciona

1. Usuário faz upload do arquivo `.pptx`
2. A app converte cada slide em imagem (via LibreOffice)
3. Cada imagem é enviada para o modelo de visão do Groq
4. As notas geradas são inseridas em cada slide
5. O `.pptx` modificado fica disponível para download

---

## Deploy no EasyPanel (Dockerfile)

1. No EasyPanel, crie um novo **App** → escolha **Dockerfile**
2. Aponte para este repositório Git
3. Porta: `5000`
4. Configure as variáveis de ambiente (veja seção abaixo)
5. Clique em **Deploy**

---

## Variáveis de ambiente

Configure na seção **Environment** do seu serviço no EasyPanel:

| Variável | Obrigatória | Descrição |
|----------|-------------|-----------|
| `GROQ_API_KEY` | Não | Se definida, oculta o campo de API key na interface — a chave é usada automaticamente |
| `SECRET_KEY` | Recomendada | Chave secreta para sessão Flask. Use um valor aleatório em produção |
| `APP_USUARIO_01` | Não | Usuário 1 (ativa o login se definido junto com `APP_SENHA_01`) |
| `APP_SENHA_01` | Não | Senha do usuário 1 |
| `APP_USUARIO_02` | Não | Usuário 2 |
| `APP_SENHA_02` | Não | Senha do usuário 2 |
| `APP_USUARIO_NN` | Não | Continue o padrão até `_99` para quantos usuários precisar |

> **Se nenhum usuário for definido**, o app funciona sem login.

### Exemplo de configuração no EasyPanel

```
GROQ_API_KEY=gsk_suachaveaqui
APP_USUARIO=admin
APP_SENHA=suasenhaforte
SECRET_KEY=mude-esta-chave-aleatoria
```

---

## Modelos suportados (Groq Vision)

Os modelos disponíveis são carregados automaticamente da API do Groq. Em caso de falha, o app usa este fallback:

| Modelo | Descrição |
|--------|-----------|
| `meta-llama/llama-4-scout-17b-16e-instruct` | Rápido e eficiente ⚡ |
| `meta-llama/llama-4-maverick-17b-128e-instruct` | Mais detalhado 🧠 |
| `llava-v1.5-7b-4096-preview` | Alternativa leve |

---

## Stack

- **Backend**: Python + Flask
- **PPTX**: python-pptx
- **Conversão de slides**: LibreOffice headless + pdf2image
- **IA**: Groq API (modelos de visão)
- **Frontend**: HTML/CSS/JS puro (sem dependências externas)
