import pandas as pd
import openpyxl as op
from PySimpleGUI import PySimpleGUI as sg
import datetime

#######################################################
def verificar_credenciais(usuario, senha):
    if usuario.strip() != '' and senha.strip() != '':
        with open('usuarios.txt', 'r') as usuarios_file, open('senhas.txt', 'r') as senhas_file:
            usuarios = [linha.strip() for linha in usuarios_file]
            senhas = [linha.strip() for linha in senhas_file]

            return usuario in usuarios and senha in senhas

def cadastrar_usuario(novo_usuario, nova_senha, senha_adm):
    if senha_adm == '123456' and novo_usuario.strip() != '' and nova_senha.strip() != '':
        try:
            with open('usuarios.txt', 'a') as usuarios_file, open('senhas.txt', 'a') as senhas_file:
                usuarios_file.write(f'\n{novo_usuario:<30}')
                senhas_file.write(f'\n{nova_senha:<30}')
            return True
        except Exception as e:
            print(f'Houve um erro ao cadastrar o usuário: {e}')
            return False
    else:
        return False

#########################################################################
pagamento = 0
total = 0
horas_not = 0
horas_ex = 0
horas_ex_mes = 0
total_mes = 0
horas_not_mes = 0
a = 0
peri = 'Não'
salario = 0
insa = 'Não'
codigo = 0

sg.theme('Black')
layout = [
    [sg.Text('DESENVOLVIMENTO AUTOMÁTICO DE HOLERITES')],
    [sg.Text('Usuário:'), sg.Input(key='Usuário')],
    [sg.Text('Senha:  '), sg.Input(key='Senha', password_char='*')],
    [sg.Button('Entrar'), sg.Button('Cadastrar novo usuário')],
    [sg.Text('Idealizado e executado por: José Otávio Junqueira Ramos')]
]

layout2 = [
    [sg.Text('')],
    [sg.Text('Funcionário nª'), sg.In(key='codigo')],
    [sg.Button('Ok')],
    [sg.Text('')],
    [sg.Text('Idealizado e executado por: José Otávio Junqueira Ramos')]
]

layout3 = [
    [sg.Text('Novo usuário:'), sg.Input(key='Novo Usuário')],
    [sg.Text('Nova senha:  '), sg.Input(key='Nova Senha')],
    [sg.Text('Senha do Administrador:'), sg.Input(key='Senha Adm', password_char='*')],
    [sg.Button('Cadastrar')]
]

janela3 = sg.Window('Cadastrar', layout3)
janela2 = sg.Window('Holerite Generator', layout2)
janela = sg.Window('Tela de login', layout)

#################################################################
while True:
    eventos, valores = janela.Read()
    if eventos == sg.WIN_CLOSED:
        break
    if eventos == 'Cadastrar novo usuário':
        janela.hide()
        while True:
            eventos_cadastro, valores_cadastro = janela3.read()
            if eventos_cadastro == sg.WIN_CLOSED:
                janela.un_hide()
                break
            if cadastrar_usuario(valores_cadastro['Novo Usuário'], valores_cadastro['Nova Senha'],
                                 valores_cadastro['Senha Adm']):
                if eventos_cadastro == 'Cadastrar':
                    sg.popup('Usuário cadastrado com sucesso!')
                else:
                    sg.popup('Falha ao cadastrar usuário!')
                janela3.close()
                janela.un_hide()
            else:
                sg.popup('Digite valores válidos ou a senha de Admnistrador corretamente!')
    if eventos == 'Entrar':
        if verificar_credenciais(valores['Usuário'], valores['Senha']):
            janela.close()
            while True:
                eventos, valores = janela2.Read()
                if eventos == sg.WIN_CLOSED:
                    break
                try:
                    codigo = int(valores['codigo'])
                except:
                    sg.popup('DIGITE UM NÚMERO VÁLIDO!')
                else:
                    pass
                if 1 <= codigo <= 7089:
                    sg.popup('Holerite está sendo gerado!\n Pressione "OK" e aguarde!')
                    try:

                        while True:
                            tabela = pd.read_excel(r'C:\Users\joseo\Desktop\Pycharm\Projeto\funcionarios.xlsx',
                                                   sheet_name=a)
                            tabela['HE'] = pd.to_numeric(tabela['HE'], errors='coerce')
                            tabela['HS'] = pd.to_numeric(tabela['HS'], errors='coerce')

                            hora_entrada = (tabela.loc[tabela['Código'] == codigo, 'HE']).item()
                            hora_saida = (tabela.loc[tabela['Código'] == codigo, 'HS']).item()
                            salario_base = tabela.loc[tabela['Código'] == codigo, 'Salário'].item()

                            # calcular horas trabalhadas
                            if hora_saida > hora_entrada:
                                total = hora_saida - hora_entrada
                            else:
                                total = (24 - hora_entrada) + hora_saida

                            # calcular horas noturnas
                            if hora_entrada <= 5:
                                horas_not = 5 - hora_entrada
                            elif hora_saida == 23:
                                horas_not = 1
                            elif hora_entrada < 13 and total == 10 or hora_entrada < 14 and total == 9 or hora_entrada < 15 and total == 8:
                                horas_not = 0
                            elif hora_entrada > hora_saida:
                                if hora_entrada <= 22:
                                    x = 22 - hora_entrada
                                if hora_saida >= 5:
                                    y = hora_saida - 5
                                else:
                                    y = 0
                                if hora_entrada == 23:
                                    horas_not = 6
                                else:
                                    horas_not = total - x - y

                            # calcular horas extras
                            if hora_saida > hora_entrada:
                                horas_ex = (hora_saida - hora_entrada) - 8
                            else:
                                horas_ex = ((24 - hora_entrada) + hora_saida) - 8

                            horas_ex_mes += horas_ex
                            total_mes += total
                            horas_not_mes += horas_not

                            a += 1
                            if a == 31:
                                break
                        # verificar periculosidade e insalubridade
                        if (tabela.loc[tabela['Código'] == codigo, 'PERICULOSIDADE']).item() == 1:
                            peri = 'Sim'
                        elif (tabela.loc[tabela['Código'] == codigo, 'INSALUBRIDADE']).item() != 0:
                            insa = (tabela.loc[tabela['Código'] == codigo, 'INSALUBRIDADE']).item()
                        else:
                            pass

                        # calcular insalubridade e periculosidade
                        salario = salario_base
                        if peri == 'Sim':
                            salario = salario_base * 1.3
                        elif insa == 1:
                            insa = 'Grau 1'
                            salario = salario_base * 1.1
                        elif insa == 2:
                            insa = 'Grau 2'
                            salario = salario_base * 1.2
                        elif insa == 3:
                            insa = 'Grau 3'
                            salario = salario_base * 1.4

                        # calcular valor horas extras
                        salario_hora = salario / 220
                        if horas_ex_mes == 0:
                            pagamento = salario
                        elif horas_ex_mes > 0:
                            pagamento = salario + (salario_hora * 1.5) * horas_ex_mes

                        # calcular valor horas noturnas
                        if horas_not_mes == 0:
                            pass
                        else:
                            pagamento += (salario_hora * 0.2) * horas_not_mes

                        # passar para o excel (holerite)
                        holerite = op.load_workbook(r'C:\Users\joseo\Desktop\Pycharm\Projeto\Holerite.xlsx')
                        aba = holerite.active
                        aba['C8'] = codigo
                        aba['G5'] = datetime.datetime.now().strftime("%B / %Y")
                        for celula in aba['D']:
                            if celula.value == 'SALARIO':
                                linha = celula.row
                                aba[f'H{linha}'] = salario_base
                            if celula.value == 'HORAS EX.':
                                linha = celula.row
                                aba[f'G{linha}'] = horas_ex_mes
                                aba[f'H{linha}'] = (salario_hora * 1.5) * horas_ex_mes
                            if celula.value == 'HORAS NOTURNAS':
                                linha = celula.row
                                aba[f'G{linha}'] = horas_not_mes
                                aba[f'H{linha}'] = (salario_hora * 0.2) * horas_not_mes
                            if celula.value == 'INSALUBRIDADE':
                                linha = celula.row
                                aba[f'G{linha}'] = insa
                                if insa == 'Grau 1':
                                    aba[f'H{linha}'] = salario_base * 0.1
                                elif insa == 'Grau 2':
                                    aba[f'H{linha}'] = salario_base * 0.2
                                elif insa == 'Grau 3':
                                    aba[f'H{linha}'] = salario_base * 0.4
                            if celula.value == 'PERICULOSIDADE':
                                linha = celula.row
                                aba[f'G{linha}'] = peri
                                if peri == 'Sim':
                                    aba[f'H{linha}'] = f'{salario_base * 0.3:.2f}'
                        holerite.save(rf'C:\Users\joseo\Desktop\Pycharm\Projeto\HOLERITES\Funcionario_nº{codigo}.xlsx')
                        # proximo funcionario
                        pagamento = 0
                        total = 0
                        horas_not = 0
                        horas_ex = 0
                        horas_ex_mes = 0
                        total_mes = 0
                        horas_not_mes = 0
                        hora_saida = 0
                        hora_entrada = 0
                        salario = 0
                        a = 0
                        peri = 'Não'
                        insa = 'Não'
                        sg.popup('Holerite gerado com sucesso!\n Encontre-o na pasta de Holerites!')
                    except Exception as e:
                        sg.popup(f'Houve um erro ao gerar o holerite: {e}')
        else:
            sg.popup('Usuário ou senha inválidos!')
