# 📊 Automated Dashboard: Python, Excel to Outlook (macOS)

Este projeto automatiza o fluxo de dados entre exportações brutas de plataformas de ensino (LMS) e relatórios executivos. Ele processa indicadores de aderência de treinamentos corporativos e gera um dashboard visual enviado nativamente via **Microsoft Outlook** no macOS.

## 🎯 O Problema

Processar manualmente relatórios de treinamento é uma tarefa repetitiva e sujeita a erros humanos. Exportações de sistemas globais geralmente vêm com termos em inglês e formatos de dados "crus", exigindo cálculos de percentuais e pivotagem de tabelas antes de serem apresentados à gestão. Além disso, automatizar o envio via scripts no macOS costuma ser complexo devido às restrições de segurança (Sandboxing) do sistema operacional.

## 💡 A Solução

A aplicação utiliza a biblioteca **Pandas** para realizar o *Data Wrangling* (limpeza, tradução e cálculo de KPIs) e a ponte **Appscript** para comandar o Microsoft Outlook nativo.

* **Tradução Automática:** Converte status técnicos (Completed, In Progress, Not Started) para uma linguagem corporativa em português.
* **Cálculo de KPI:** Gera automaticamente a taxa de aderência percentual consolidada por curso.
* **UX Premium:** O e-mail é entregue com um design moderno baseado em cartões (HTML/CSS), facilitando a leitura em dispositivos móveis e desktops.
* **Segurança Nativa:** Ao utilizar o `appscript`, o projeto contorna a necessidade de armazenar senhas ou tokens de e-mail no código, utilizando a própria sessão autenticada do usuário.

## 🛠️ Tecnologias Utilizadas

* **Python 3.9+**
* **Pandas**: Processamento e análise de dados matriciais.
* **Appscript**: Automação de aplicativos nativos do macOS (AppleScript bridge).
* **Mactypes**: Gerenciamento de permissões de arquivos (Alias) para o macOS.

## 🚀 Como Utilizar

1. **Pré-requisitos:** Certifique-se de ter o Microsoft Outlook instalado e configurado no seu Mac.

2. **Instalação:**
```bash
pip install pandas openpyxl appscript

```

3. **Configuração:** No bloco principal do script, aponte o caminho do seu arquivo Excel e o e-mail do destinatário.
4. **Execução:**
```bash
python email_automatico_outlook.py

```


## 🔒 Segurança e Boas Práticas

* **Zero Auth Exposure:** O código não solicita nem armazena credenciais.
* **Sandboxing Compliance:** Utiliza `mactypes.Alias` para garantir que o Outlook tenha permissão de leitura sobre o anexo, evitando o erro comum `OSERROR -2700`.
* **Clean Code:** Métodos com responsabilidade única e variáveis centralizadas para fácil manutenção.

---

**Desenvolvido para otimizar processos de report de treinamento.** 📈
