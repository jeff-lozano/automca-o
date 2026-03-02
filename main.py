from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Controle Financeiro"

# Cabeçalhos
ws.append(["Tipo", "Descrição", "Valor"])

receitas = 0
despesas = 0

while True:
    tipo = input("Digite o tipo (R para Receita, D para Despesa ou S para sair): ").upper()
    
    if tipo == "S":
        break
    
    descricao = input("Descrição: ")
    valor = float(input("Valor: "))

    ws.append([tipo, descricao, valor])

    if tipo == "R":
        receitas += valor
    elif tipo == "D":
        despesas += valor

saldo = receitas - despesas

# Resumo final
ws.append([])
ws.append(["Total Receitas", receitas])
ws.append(["Total Despesas", despesas])
ws.append(["Saldo Final", saldo])

wb.save("controle_financeiro.xlsx")

print("Planilha criada com sucesso!")
