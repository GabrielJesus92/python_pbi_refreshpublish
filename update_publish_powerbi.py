from pywinauto import Application
import pyautogui
import time
import subprocess

# Caminho para o arquivo PBIX
pbix_path = r'C:/Users/gblsj/OneDrive/Documentos/PBI/Dashboard de Performance - Taylor Swift.pbix'
print(f"Path do arquivo PBIX definido: {pbix_path}")

# Caminho para o executável do Power BI Desktop instalado via Microsoft Store
powerbi_path = r'C:/Program Files/Microsoft Power BI Desktop/bin/PBIDesktop.exe'
print(f"Path do instalador do PBI definido: {powerbi_path}")

# Comando para abrir o Power BI Desktop com o arquivo PBIX
command = f'"{powerbi_path}" "{pbix_path}"'

# Use subprocess.Popen para executar o comando de forma assíncrona
process = subprocess.Popen(command, shell=True)
print("Arquivo PBIX aberto com sucesso!")

# Aguarde a janela do Power BI Desktop abrir
time.sleep(20)  # Ajuste o tempo conforme necessário para garantir que o arquivo PBIX seja carregado
print("Tempo de espera para iniciar a manipulação do arquivo concluído")

# Conecte-se à aplicação do Power BI Desktop
app = Application().connect(title_re='.*Dashboard de Performance*')
pbi_window = app.window(title_re='.*Dashboard de Performance*')
pbi_window.set_focus()
print("Janela encontrada com sucesso")

# Clique na barra de título para garantir que o foco esteja correto
# pbi_window.click_input(double=True)  # Clica duas vezes para focar a janela

# Execute o refresh
pyautogui.hotkey('alt')  # Abra o menu de refresh (atalho pode variar)
time.sleep(2)
pyautogui.hotkey('c')  # Abra o menu de refresh (atalho pode variar)
time.sleep(2)
pyautogui.press('r')  # Pressione 'r' para refresh (atalho pode variar)

# Aguarde a conclusão do refresh
time.sleep(30)  # Ajuste o tempo conforme necessário para garantir que o refresh seja concluído
# Clique na barra de título para garantir que o foco esteja correto
pbi_window.click_input(double=True)  # Clica duas vezes para focar a janela

# Execute o save
pyautogui.hotkey('alt')  # Abra o menu de refresh (atalho pode variar)
time.sleep(2)
pyautogui.hotkey('1')  # Abra o menu de refresh (atalho pode variar)

# Aguarde a conclusão do save
time.sleep(10)  # Ajuste o tempo conforme necessário para garantir que o refresh seja concluído
# Clique na barra de título para garantir que o foco esteja correto
pbi_window.click_input(double=True)  # Clica duas vezes para focar a janela

# Publicar no workspace
pyautogui.hotkey('alt')  # Abra o menu de refresh (atalho pode variar)
time.sleep(2)
pyautogui.hotkey('c')  # Abra o menu de refresh (atalho pode variar)
time.sleep(2)
pyautogui.press('p')  # Pressione 'r' para refresh (atalho pode variar)
time.sleep(5)
pyautogui.press('tab')  # Pressione 'r' para refresh (atalho pode variar)
time.sleep(1)
pyautogui.press('tab')  # Pressione 'r' para refresh (atalho pode variar)
time.sleep(1)
pyautogui.press('tab')  # Pressione 'r' para refresh (atalho pode variar)
time.sleep(2)
pyautogui.press('down')  # Pressione 'r' para refresh (atalho pode variar)
time.sleep(2)
pyautogui.press('tab')  # Pressione 'r' para refresh (atalho pode variar)
time.sleep(2)
pyautogui.press('enter')  # Pressione 'r' para refresh (atalho pode variar)
time.sleep(5)
pyautogui.press('enter')  # Pressione 'r' para refresh (atalho pode variar)
# Aguarde a conclusão da publicação
time.sleep(60)
# fecha a popup da conclusão da publicação
pyautogui.press('tab')  # Pressione 'r' para refresh (atalho pode variar)
time.sleep(2)
pyautogui.press('enter')  # Pressione 'r' para refresh (atalho pode variar)


# pausa após fechar a popup
time.sleep(3)  # Ajuste o tempo conforme necessário para garantir que a publicação seja concluída

# Feche o Power BI Desktop
pbi_window.close()

print("Processo de automação concluído.")
