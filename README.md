# MonitorMicroSIP

Automatize o controle do MicroSIP com base no bloqueio e desbloqueio da sessão do Windows.

## 💡 Objetivo

Este serviço foi criado para **encerrar o MicroSIP quando a sessão do usuário for bloqueada** e **reiniciá-lo automaticamente ao desbloquear**. Ele utiliza eventos do sistema operacional para garantir desempenho e resposta rápida, sem verificações contínuas de processo.

---

## ⚙️ Funcionalidades

- 🔐 Encerra automaticamente o `microsip.exe` ao bloquear a sessão do Windows.
- 🔓 Inicia automaticamente o `microsip.exe` ao desbloquear a sessão.
- 🧠 Usa a API nativa do Windows para receber eventos de sessão em tempo real.
- 📄 Gera logs rotativos semanais em `C:\MicroSIPService\servico.log`.
- 🚫 **Sem verificação contínua de processos**, evitando atrasos e uso excessivo de CPU.

---

## 🛠️ Requisitos

- **Sistema Operacional:** Windows 10 ou superior
- **Python:** Recomendado Python 3.10 ou 3.12  
- **Pacotes necessários:**  
  - `pywin32`

Instale com:

```bash
pip install pywin32
```

---

## 📂 Estrutura dos arquivos

```
MonitorMicroSIP_v2.pyw         # Código principal do monitor
C:\MicroSIPService\servico.log # Log rotativo semanal do serviço
```

---

## 🚀 Como usar

1. **Configure o caminho do MicroSIP (já automático via USERPROFILE):**
   - O script busca em:  
     `%USERPROFILE%\AppData\Local\MicroSIP\microsip.exe`

2. **Execute como script invisível (modo produção):**

```bash
pythonw MonitorMicroSIP_v2.pyw
```

3. **Ou compile como `.exe`:**

```bash
pyinstaller --noconsole --onefile MonitorMicroSIP_v2.pyw
```

(Opcional: adicione `.spec` com ícone ou configurações adicionais)

---

## 🪵 Log

O log de execução será salvo em:
```
C:\MicroSIPService\servico.log
```
Contém eventos como:
- Início/encerramento do MicroSIP
- Eventos de bloqueio/desbloqueio
- Erros de sistema ou execução

---

## 🔄 Histórico de versões

[1.0.0] - 2025-07-01

Lançamento Inicial

Adicionado

✨ Monitoramento de eventos de sessão: Implementado monitoramento de bloqueio (WTS_SESSION_LOCK) e desbloqueio (WTS_SESSION_UNLOCK) de sessão no Windows usando a API Wtsapi32.
✨ Gerenciamento de processos: Adicionada funcionalidade para iniciar e encerrar um programa selecionado pelo usuário com base em eventos de sessão.
✨ Verificação de processos: Utiliza o comando tasklist para verificar se o programa está em execução antes de iniciá-lo ou encerrá-lo.
✨ Interface de seleção de programa: Integração com tkinter para permitir que o usuário selecione o programa a ser monitorado via interface gráfica.
✨ Tratamento de erros: Adicionado tratamento básico de erros para verificação de processos, execução de programas e loop de mensagens do Windows.

Notas
Esta é a primeira versão estável do programa, projetada para automatizar o controle de programas com base em eventos de bloqueio e desbloqueio de sessão no Windows.

O programa depende do comando tasklist para verificação de processos e da API Wtsapi32 para monitoramento de eventos.


[2.0.0] - 2025-07-18

Alterações Significativas

Removido
🔥 Verificação com tasklist: Removida a função de verificação de processos usando o comando tasklist, eliminando a dependência de verificações externas para determinar se o programa está em execução.
🔥 Interface de seleção de programa: Removida a integração com tkinter para seleção manual do programa, substituída por um caminho fixo para o MicroSIP.

Melhorado
🧠 Desempenho no controle do MicroSIP: Otimizada a inicialização do MicroSIP ao usar um caminho fixo baseado em USERPROFILE, reduzindo latência e simplificando o processo de execução.
🧠 Estabilidade: Aprimorado o tratamento de erros com mensagens de log mais detalhadas e verificações adicionais (ex.: validação do caminho do programa).
🧠 Logging: Adicionado sistema de logging robusto com TimedRotatingFileHandler, salvando logs em C:\MicroSIPService\servico.log com rotação semanal.

Adicionado

✅ Controle baseado em eventos: Implementado gerenciamento 100% baseado em eventos do Windows, eliminando verificações periódicas de processos e utilizando a API Wtsapi32 para iniciar e encerrar o MicroSIP de forma mais eficiente.
✅ Caminho fixo para o MicroSIP: O programa agora usa um caminho predefinido (%USERPROFILE%\AppData\Local\MicroSIP\microsip.exe) para gerenciar o MicroSIP, eliminando a necessidade de seleção manual.
✅ Criação de diretório de logs: Adicionada criação automática do diretório C:\MicroSIPService para armazenamento de logs.
---

## 📄 Licença

Este projeto é proprietário e pode ser adaptado internamente de acordo com as necessidades da empresa. Consulte o autor/responsável pela automação para ajustes.

---

## 👨‍💻 Autor

William Carvalho  
Desenvolvedor Interno – Green Benefícios
