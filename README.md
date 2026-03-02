# Active Directory Password Reset Automation with AI

Automação completa de redefinição de senha no Active Directory baseada em solicitações recebidas por email no Outlook, utilizando Inteligência Artificial para identificar o usuário automaticamente.

Este projeto foi desenvolvido para otimizar o fluxo de trabalho de equipes de Service Desk e Suporte de TI, eliminando tarefas manuais e reduzindo o tempo de atendimento.

---

## Funcionalidades

- Leitura automática de emails do Outlook
- Identificação do nome do usuário utilizando IA (Ollama + Llama3)
- Busca automática do usuário no Active Directory
- Redefinição automática de senha
- Desbloqueio automático da conta
- Resposta automática ao email com as credenciais
- Interface simples via terminal

---

## Tecnologias utilizadas

- Python 3
- Outlook COM API (win32com)
- PowerShell
- Active Directory
- Ollama (LLM local)
- Llama 3
- JSON

---

## Arquitetura do fluxo
Email recebido → Outlook → Python Script → IA extrai nome → Busca no AD →
Reset de senha → Desbloqueio → Resposta automática ao email
