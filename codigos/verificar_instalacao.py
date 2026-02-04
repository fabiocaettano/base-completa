import sys

print("=" * 50)
print("VERIFICAÇÃO DE INSTALAÇÃO PYTHON")
print("=" * 50)

# Versão do Python
print(f"Python version: {sys.version}")
print(f"Python version info: {sys.version_info}")

# Verificar bibliotecas
bibliotecas = ['pandas', 'openpyxl', 'numpy', 'json', 'datetime', 'collections']

for biblioteca in bibliotecas:
    try:
        __import__(biblioteca)
        print(f"✓ {biblioteca}: INSTALADA")
    except ImportError:
        print(f"✗ {biblioteca}: NÃO INSTALADA")

print("=" * 50)