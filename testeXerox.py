import pandas as pd
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 3.2                 

def abrirGoogle():
    pyautogui.press("win")
    pyautogui.write("chrome")
    pyautogui.press("enter")

def entrarNoLink(link):
    pyperclip.copy(link)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press("enter")

abrirGoogle()

# Seleciona o meu perfil
pyautogui.press('tab', presses=4, interval=0.01)
pyautogui.press("enter", interval=0.01)

# Entra no link do arquivo excel e instala ele
entrarNoLink("https://drive.google.com/drive/folders/1B9MivZ5LxxHVSqiAas5epbwaFbcd4o7f?usp=sharing")
time.sleep(3)

pyautogui.press('tab', presses=2, interval=0.01)
pyautogui.press('enter', interval=0.1)
pyautogui.press('down')
pyautogui.hotkey('enter', interval=0.1)
pyautogui.press('left', presses=2, interval=0.1)
pyautogui.press("enter", presses=2, interval=0.01)

time.sleep(5)
pyautogui.hotkey('alt', 'f4', interval=0.1)


pyautogui.press("win")
pyautogui.write("explorador de arquivos")
pyautogui.press("right")
pyautogui.press("down", presses=3, interval=0.1)
pyautogui.press("enter")

# Carregar o arquivo Excel
file_path = 'C:\\Users\\garot\\Downloads\\Vendas.xlsx'
xls = pd.ExcelFile(file_path)

# Carregar a folha "Plan1"
df = pd.read_excel(file_path, sheet_name='Plan1')

# Calcular estatísticas para 'Valor Unitário' e 'Valor Final'
stats = {
    'Estatística': ['Média', 'Mediana', 'Desvio Padrão', 'Valor Mínimo', 'Valor Máximo', 'Soma'],
    'Valor Unitário': [
        df['Valor Unitário'].mean(),
        df['Valor Unitário'].median(),
        df['Valor Unitário'].std(),
        df['Valor Unitário'].min(),
        df['Valor Unitário'].max(),
        df['Valor Unitário'].sum()
    ],
    'Valor Final': [
        df['Valor Final'].mean(),
        df['Valor Final'].median(),
        df['Valor Final'].std(),
        df['Valor Final'].min(),
        df['Valor Final'].max(),
        df['Valor Final'].sum()
    ]
}

# Criar um DataFrame com as estatísticas
stats_df = pd.DataFrame(stats)

# Escrever os dados originais e as estatísticas em um novo arquivo Excel
output_file_path = 'C:\\Users\\garot\\Downloads\\Vendas com Estatísticas.xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    df.to_excel(writer, sheet_name='Plan1', index=False)
    stats_df.to_excel(writer, sheet_name='Estatísticas', index=False)

print(f'Arquivo salvo em: {output_file_path}')    

pyautogui.hotkey('f5', interval=0.1)

pyautogui.press("win")
pyperclip.copy("Vendas com Estatísticas")  
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("enter")

time.sleep(5)

pyautogui.hotkey('ctrl', 'shift', 'pagedown', interval=0.1)


abrirGoogle()
# Seleciona o meu perfil
pyautogui.press('tab', presses=4, interval=0.01)
pyautogui.press("enter", interval=0.01)

entrarNoLink("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox?compose=new")
time.sleep(3)



pyautogui.write("Tiago.Silva@xerox.com")  # Preencher o destinatário
pyautogui.press("tab")  # Selecionar o e-mail
pyautogui.press("tab")  # Passar para o campo de assunto

assunto = "Novo Relatório de Vendas de Dezembro com Estatísticas Detalhadas"
pyperclip.copy(assunto)  # Preencher o assunto
pyautogui.hotkey('ctrl', 'v')
pyautogui.press("tab")  # Passar para o campo de corpo

texto = f"""
Olá Equipe,

Estou animado para compartilhar com vocês o novo relatório de vendas de dezembro, agora com um toque especial! Além das informações habituais, adicionamos estatísticas detalhadas para as colunas "Valor Unitário" e "Valor Final" que ajudarão a obter insights valiosos sobre nosso desempenho.

### Destaques do Relatório:

Na planilha "Plan1", além dos dados de vendas, vocês encontrarão as seguintes estatísticas:

- **Média**: O valor médio dos itens vendidos.
- **Mediana**: O valor central dos itens vendidos.
- **Desvio Padrão**: A medida da dispersão dos valores.
- **Valor Mínimo**: O menor valor registrado.
- **Valor Máximo**: O maior valor registrado.
- **Soma**: A soma total dos valores.

#### Estatísticas Calculadas:
| Estatística    | Valor Unitário | Valor Final |
|----------------|----------------|-------------|
| **Média**          | {df['Valor Unitário'].mean():.2f}         | {df['Valor Final'].mean():.2f}      |
| **Mediana**        | {df['Valor Unitário'].median():.2f}       | {df['Valor Final'].median():.2f}    |
| **Desvio Padrão**  | {df['Valor Unitário'].std():.2f}          | {df['Valor Final'].std():.2f}       |
| **Valor Mínimo**   | {df['Valor Unitário'].min():.2f}          | {df['Valor Final'].min():.2f}       |
| **Valor Máximo**   | {df['Valor Unitário'].max():.2f}          | {df['Valor Final'].max():.2f}       |
| **Soma**           | {df['Valor Unitário'].sum():.2f}          | {df['Valor Final'].sum():.2f}       |

Obrigado e fiquem à vontade para entrar em contato se tiverem alguma dúvida.

Atenciosamente,
Tiago
"""

pyperclip.copy(texto)  # Preencher o corpo do e-mail
pyautogui.hotkey('ctrl', 'v')

# Opcional: enviar o e-mail
# pyautogui.hotkey('ctrl', 'enter')

pyautogui.press('tab', presses=3, interval=0.01)
pyautogui.press('enter', interval=0.01)

pyperclip.copy("Vendas com Estatísticas")  #Nome da pasta que vai ser usada
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter', interval=0.01)
pyautogui.press('tab', presses=2, interval=0.01)
pyautogui.press('enter', interval=0.01)

