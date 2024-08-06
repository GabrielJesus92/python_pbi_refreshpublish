# Documentação da Automação do Power BI Desktop

Este script Python automatiza a abertura de um arquivo PBIX no Power BI Desktop, realiza um refresh dos dados, salva o arquivo e publica no workspace. Utiliza as bibliotecas `pywinauto` para controle da interface do usuário e `pyautogui` para simular a interação com a interface do usuário.

## Funcionalidades

1. **Abrir Power BI Desktop**: Inicia o Power BI Desktop e carrega um arquivo PBIX especificado.
2. **Atualizar Dados**: Executa a atualização dos dados no Power BI Desktop.
3. **Salvar Arquivo**: Salva o arquivo PBIX após a atualização.
4. **Publicar no Workspace**: Publica o arquivo PBIX em um workspace no Power BI.
5. **Fechar o Power BI Desktop**: Fecha a aplicação Power BI Desktop após a conclusão da automação.

## Requisitos

- **Bibliotecas Python**: `pywinauto`, `pyautogui`, `time`, `subprocess`
- **Caminho do Arquivo PBIX**: Defina o caminho para o arquivo PBIX no seu sistema.
- **Caminho do Executável do Power BI Desktop**: Defina o caminho para o executável do Power BI Desktop instalado via Microsoft Store.

## Estrutura do Código

### Etapas de Automação:

1. **Definição de Caminhos**:
   - **Arquivo PBIX**: `pbix_path`
   - **Executável do Power BI Desktop**: `powerbi_path`

2. **Abrir Power BI Desktop**:
   - Usa `subprocess.Popen` para abrir o arquivo PBIX no Power BI Desktop.

3. **Aguardar Abertura**:
   - Aguarda 20 segundos para garantir que o Power BI Desktop e o arquivo PBIX estejam completamente carregados.

4. **Conectar à Aplicação Power BI**:
   - Usa `pywinauto` para se conectar à janela do Power BI Desktop e definir o foco na janela do arquivo PBIX.

5. **Atualizar Dados**:
   - Usa `pyautogui` para simular a sequência de teclas para abrir o menu de refresh e executar a atualização dos dados.
   - Aguarda 30 segundos para garantir que o refresh seja concluído.

6. **Salvar Arquivo**:
   - Usa `pyautogui` para simular a sequência de teclas para salvar o arquivo PBIX.
   - Aguarda 10 segundos para garantir que o save seja concluído.

7. **Publicar no Workspace**:
   - Usa `pyautogui` para simular a sequência de teclas para abrir o menu de publicação e concluir a publicação no workspace.
   - Aguarda 60 segundos para garantir que a publicação seja concluída e fecha a popup de conclusão.

8. **Fechar Power BI Desktop**:
   - Fecha a aplicação Power BI Desktop.

## Uso

1. **Preparação**:
   - Certifique-se de que o caminho para o arquivo PBIX e o executável do Power BI Desktop estão corretos.
   
2. **Executar o Script**:
   ```bash
   python nome_do_arquivo.py

# Contribuições
Sinta-se à vontade para melhorar o script e ajustar os tempos e atalhos conforme necessário. Abra uma issue ou um pull request para colaborar.
