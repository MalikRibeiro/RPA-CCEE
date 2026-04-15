# RPA CCEE

Automação semi-automática para extração de relatórios em PDF do portal da CCEE.

## Visão geral

O projeto contém duas versões principais:

- `v1/`: versão original com `iniciar_robo.bat` e conexão manual ao Chrome em modo debug.
- `v2/`: versão mais recente que inicia o Chrome automaticamente em modo debug e inicia o fluxo pelo próprio script.

Ambas as versões utilizam Selenium e PyAutoGUI para controlar o Chrome e salvar relatórios em PDF.

## Estrutura do projeto

- `JavaScript/` - scripts auxiliares e versões anteriores em JavaScript.
- `v1/` - versão 1 do robô com arquivo `iniciar_robo.bat`.
- `v2/` - versão 2 do robô com inicialização automática do Chrome.

## Requisitos

- Python 3.x instalado.
- Google Chrome instalado.
- Dependências Python:
  - `selenium`
  - `pyautogui`

> Observação: `v2/requirements.txt` não lista todas as dependências necessárias. Use `pip install selenium pyautogui` se precisar.

## Como configurar

Antes de executar, edite o arquivo de configuração da versão que você irá usar:

- `v1/config.py`
- `v2/config.py`

Ajuste pelo menos os campos:

- `MES` - mês do relatório.
- `ANO` - ano do relatório.
- `NOME_RELATORIO` - código do relatório.
- `PASTA_DOWNLOAD` - pasta onde os PDFs serão salvos.
- `empresas_alvo` - lista de empresas a serem processadas.

## Execução

### Usando `v1`

1. Feche todas as instâncias do Google Chrome.
2. Execute `v1/iniciar_robo.bat`.
3. No Chrome que abriu, faça login na CCEE e navegue até o relatório desejado.
4. Abra um terminal ou VS Code em `v1/`.
5. Execute:

```bash
python app.py
```

### Usando `v2`

1. Abra um terminal em `v2/`.
2. Execute:

```bash
python app.py
```

3. O script abrirá o Chrome em modo debug e pedirá para você fazer login.
4. Após preparar o relatório, confirme no terminal e o robô continuará a extração.

## Observações importantes

- Não mova o mouse nem use o teclado enquanto o robô estiver salvando o PDF.
- Mantenha o Chrome aberto durante toda a execução.
- Os arquivos gerados serão salvos em `PASTA_DOWNLOAD` conforme definido em `config.py`.

## Dicas rápidas

- Se o robô não encontrar a página certa, verifique `URL_ESPERADA_CCEE` em `config.py`.
- Se o salvamento de PDF falhar, talvez seja necessário ajustar a posição ou a janela do navegador.
- Sempre use a versão apropriada para o fluxo que você deseja: `v1` para a versão com .bat, `v2` para a versão com inicialização automática do Chrome.

