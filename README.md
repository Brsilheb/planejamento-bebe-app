# Planejamento do Bebê — App (Empacotamento .EXE)

Este pacote permite gerar um **executável (.exe)** do app Streamlit.

## Estrutura esperada na pasta
- `app_v2.py` (o app principal Streamlit)
- `Planejamento_Bebe_Heber.xlsx` (planilha base)
- `launcher.py` (inicializador usado pelo executável)
- `requirements.txt`
- `build_exe.bat` (atalho de build)
- `app_v2.spec` (alternativa avançada de build)

## Passo a passo (Windows)
1. **(Opcional, recomendado)** criar e ativar venv:
```
python -m venv .venv
.venv\Scriptsctivate
```
2. **Instalar dependências** e **gerar .exe**:
```
pip install --upgrade pip
pip install -r requirements.txt
pyinstaller --onefile --noconsole ^
  --add-data "app_v2.py;." ^
  --add-data "Planejamento_Bebe_Heber.xlsx;." ^
  launcher.py
```
3. **Executar**: abra `dist\launcher.exe` (o navegador abrirá o app).

> Dica: use `build_exe.bat` para automatizar os comandos acima.

### Observações
- Se quiser **alterar o nome do .exe**, troque `launcher.py` por outro nome e ajuste o comando do PyInstaller.
- Se for **rodar em outra máquina**, inclua o `Planejamento_Bebe_Heber.xlsx` ao lado do `.exe` ou já embuta com `--add-data` (como acima).
- Caso seu antivírus alerte, marque o `launcher.exe` como confiável.
- Porta padrão: `8501`. Para mudar, defina `STREAMLIT_SERVER_PORT` no ambiente antes de abrir o `.exe`.
