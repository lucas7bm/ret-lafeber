import PySimpleGUI as sg
from time import sleep
import xlsxwriter

def update_counters(efd_array):
    # Definição dos contadores
    counter_0200 = 0
    counter_0205 = 0
    counter_0206 = 0
    counter_0210 = 0
    counter_0220 = 0
    counter_0460 = 0
    counter_0990 = 0
    counter_c195 = 0
    counter_c170 = 0
    counter_c197 = 0
    counter_C990 = 0
    counter_1990 = 0
    counter_9900 = 0
    counter_9990 = 0
    counter_9999 = 0
    # Contagem
    for line in efd_array:
        counter_9999 += 1
        if line[0][0] == '0':
            counter_0990 += 1
        if line[0][0] == '1':
            counter_1990 += 1
        if line[0][0] == 'C':
            counter_C990 += 1
        if line[0][0] == '9':
            counter_9990 += 1
        if line[0] == '0200':
            counter_0200 += 1
        if line[0] == '0205':
            counter_0205 += 1
        if line[0] == '0206':
            counter_0206 += 1
        if line[0] == '0210':
            counter_0210 += 1
        if line[0] == '0220':
            counter_0220 += 1
        if line[0] == '0460':
            counter_0460 += 1
        if line[0] == 'C170':
            counter_c170 += 1
        if line[0] == 'C195':
            counter_c195 += 1
        if line[0] == 'C197':
            counter_c197 += 1
        if line[0] == '9900':
            counter_9900 += 1
    # Alocando os contadores
    updated = False
    for line in efd_array:
        if line[0] == '0990':
            old_counter = line[1]
            line[1] = str(counter_0990)
            if old_counter != line[1]:
                print("Registro contador 0990 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
        if line[0] == 'C990':
            old_counter = line[1]
            line[1] = str(counter_C990)
            if old_counter != line[1]:
                print("Registro contador C990 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
        if line[0] == '1990':
            old_counter = line[1]
            line[1] = str(counter_1990)
            if old_counter != line[1]:
                print("Registro contador 1990 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
        elif line[0] == '9990':
            old_counter = line[1]
            line[1] = str(counter_9990)
            if old_counter != line[1]:
                print("Registro contador 9990 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0200':
            old_counter = line[2]
            line[2] = str(counter_0200)
            if old_counter != line[2]:
                print("Registro contador 9900|0200 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0205':
            old_counter = line[2]
            line[2] = str(counter_0205)
            if old_counter != line[2]:
                print("Registro contador 9900|0205 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0206':
            old_counter = line[2]
            line[2] = str(counter_0206)
            if old_counter != line[2]:
                print("Registro contador 9900|0206 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0210':
            old_counter = line[2]
            line[2] = str(counter_0210)
            if old_counter != line[2]:
                print("Registro contador 9900|0210 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0220':
            old_counter = line[2]
            line[2] = str(counter_0220)
            if old_counter != line[2]:
                print("Registro contador 9900|0220 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '0460':
            old_counter = line[2]
            line[2] = str(counter_0460)
            if old_counter != line[2]:
                print("Registro contador 9900|0460 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == 'C170':
            old_counter = line[2]
            line[2] = str(counter_c170)
            if old_counter != line[2]:
                print("Registro contador 9900|C170 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == 'C195':
            old_counter = line[2]
            line[2] = str(counter_c195)
            if old_counter != line[2]:
                print("Registro contador 9900|C195 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == 'C197':
            old_counter = line[2]
            line[2] = str(counter_c197)
            if old_counter != line[2]:
                print("Registro contador 9900|C197 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9900' and line[1] == '9900':
            old_counter = line[2]
            line[2] = str(counter_9900)
            if old_counter != line[2]:
                print("Registro contador 9900|9900 atualizado de " +
                    old_counter + " para " + line[2] + ".")
                updated = True
        elif line[0] == '9999':
            old_counter = line[1]
            line[1] = str(counter_9999)
            if old_counter != line[1]:
                print("Registro contador 9999 atualizado de " +
                    old_counter + " para " + line[1] + ".")
                updated = True
    if updated:
        sg.popup("Os registros contadores foram atualizados!", title="Atenção!")
    return efd_array

def fix_unusedItems(efd_array):
    new_sped = []
    wannafix = False
    # Gerando lista de items referenciados
    referenced_items = []
    for line in efd_array:
        if line[0] == "C170":
            referenced_items.append(line[2])
        if line[0] == "C197":
            referenced_items.append(line[3])
        if line[0] == "C425":
            referenced_items.append(line[1])
        if line[0] == "H010":
            referenced_items.append(line[1])
        if line[0] == "K200":
            referenced_items.append(line[2])
    for line in efd_array:
        if (line[0] == "0200" and line[1] not in referenced_items):
            ans = sg.popup_yes_no("Foram encontrados itens não referenciados em nenhum registro. Deseja corrigir?", title="Atenção!")
            if ans == "Yes":
                wannafix = True
            break
    # Removendo itens não utilizados
    if wannafix:
        print("\n\nCORREÇÕES DE ITENS NÃO REFERENCIADOS:")
        i = 0
        while i < len(efd_array):
            if efd_array[i][0] == "0200":
                if efd_array[i][1] in referenced_items:
                    new_sped.append(efd_array[i])
                else:
                    print("O item " + efd_array[i][1] + " foi excluído.")
                    children = ['0205', '0206', '0210', '0220']
                    for row in efd_array[i+1:i+5]:
                        if row[0] == '0200':
                            break
                        if row[0] in children:
                            print("Registro " + row[0] + " excluído.")
                            i+=1
            else:
                new_sped.append(efd_array[i])
            i+=1
        print("\n")
        return new_sped
    else:
        return efd_array

def get_value(str):
    try:
        value = round(float(str.replace(",", ".")), 2)
        return value
    except:
        return "err"

def set_value(value):
    try:
        return str(round(value, 2)).replace(".", ",")
    except:
        return "err"

def insert_9900_counter(efd_array, reg):
    # Caso o contador do registro 9900 não exista, o inserimos
    has_counter = False
    for row in efd_array:
        if row[0] == '9900' and row[1] == reg:
            has_counter = True
    if has_counter:
        return efd_array
    for i in range(len(efd_array)):
        if efd_array[i][0] == '9900' and efd_array[i][1] > reg:
            efd_array = efd_array[:i] + [['9900', reg, '1']] + efd_array[i:]
            break
    return efd_array

def insert_0460(efd_array, obs):
    for i in range(len(efd_array)):
        if efd_array[i][0] > '0460':
            efd_array = efd_array[:i] + [obs] + efd_array[i:]
            print("Ajuste adicionado registro 0460: " + obs[2])
            break
    efd_array = insert_9900_counter(efd_array, '0460')
    return efd_array

def create_adj_block(efd_array, nfe_key):
    adj_block = [['C195', 'estD', 'Estorno de débito - e-PTA n. 45.000023237-88 - Sub-apuração']]
    for row in efd_array:
        if row[0] == 'C100':
            current_doc = row[8]
        if row[0] == 'C170' and current_doc == nfe_key:
            COD_ITEM = row[2]
            BC_ICMS = row[12]
            ALIQ_ICMS = row[13]
            VL_ICMS = row[14]
            if get_value(VL_ICMS) == 0.00:
                continue
            adj_C197 = ['C197', 'MG23000999', 'Estorno de débito - e-PTA n. 45.000023237-88 - Sub-apuração', COD_ITEM, BC_ICMS, ALIQ_ICMS, VL_ICMS, '']
            adj_block.append(adj_C197)
    if len(adj_block) == 1:
        return []
    return adj_block

def insert_adj_block(efd_array, nfe_key, adj_block):
    current_doc = ""
    for i in range(len(efd_array)):
        if current_doc == nfe_key and ((efd_array[i][0] > 'C190') or (efd_array[i][0] == 'C100')):
            efd_array = efd_array[:i] + adj_block + efd_array[i:]
            break
        if efd_array[i][0] == 'C100':
            current_doc = efd_array[i][8]
    efd_array = insert_9900_counter(insert_9900_counter(efd_array, 'C195'), 'C197')
    return efd_array

def ajustes_RET(efd_array):
    obs_estD = ['0460', 'estD', 'Estorno de débito - e-PTA n. 45.000023237-88 - Sub-apuração']
    efd_array = insert_0460(efd_array, obs_estD)
    outgoing_invoices = []
    for row in efd_array:
        if row[0] == 'C100' and row[1] == '1':
            current_doc = row[8]
            outgoing_invoices.append(current_doc)
        else:
            continue
    for nfe_key in outgoing_invoices:
        adj_block = create_adj_block(efd_array, nfe_key)
        efd_array = insert_adj_block(efd_array, nfe_key, adj_block)
    #Colocamos o total de débitos no total de ajustes
    for row in efd_array:
        if row[0] == 'E110':
            row[6] = row[1]
            break
    return efd_array

def get_participant_data(efd_array, participant_id):
    # returns IE, CNPJ, NAME
    for row in efd_array:
        if row[0] == '0150' and row[1] == participant_id:
            return row[6], row[4], row[2]
    return "", "", ""

def subA_RET(efd_array, OUTPUT_FOLDER):
    outputsheet_rows = []
    VL_TOTAL_DEB = 0.00
    VL_TOTAL_RET = 0.00
    nfe_num = ''
    nfe_key = ''
    participant_id = ''
    VALID_CFOPS = ['101', '401', '107', ]
    for row in efd_array:
        # Somente consideraremos notas de saída
        if row[0] == 'C100':
            if row[1] == '1':
                participant_id = row[3]
                nfe_num = row[7]
                nfe_key = row[8]
            else:
                nfe_key = ''
                participant_id = ''
                nfe_num = ''
        if nfe_key and (row[0] == 'C190'):
            part_IE, part_CNPJ, part_NAME = get_participant_data(efd_array, participant_id)
            # Se o participante possui IE, geralmente é contribuinte. Consideramos-no como tal.
            # A alíquota será de 1% para todos os CFOPs válidos.
            cfop = row[2]
            vl_opr = get_value(row[4])
            vl_bc_icms = row[5]
            vl_icms = row[6]
            VL_TOTAL_DEB += get_value(vl_icms)
            if part_IE:
                if cfop[1:] in VALID_CFOPS:
                    vl_ret = vl_opr * 0.01
                    VL_TOTAL_RET += vl_ret
                    # The output row.
                    # |num|CHV|CNPJ|nome|Contrib|CFOP|VL_OP|BC|ICMS|ALIQ_RET|RET|
                    row = [nfe_num, nfe_key, part_CNPJ, part_NAME, "SIM", cfop, vl_opr, vl_bc_icms, vl_icms, "1%", vl_ret]
                    outputsheet_rows.append(row)
            # Caso o participante não seja contribuinte, a alíquota será outra.
            else:
                if cfop[1:] in VALID_CFOPS:
                    # Alíquota de 6% em operações internas
                    if cfop[0] == '5':
                        vl_ret = vl_opr * 0.06
                        VL_TOTAL_RET += vl_ret
                        row = [nfe_num, nfe_key, part_CNPJ, part_NAME, "NÃO", cfop, vl_opr, vl_bc_icms, vl_icms, "6%", vl_ret]
                    # Alíquota de 2% para operações interestaduais
                    if cfop[0] == '6':
                        vl_ret = vl_opr * 0.02
                        VL_TOTAL_RET += vl_ret
                        row = [nfe_num, nfe_key, part_CNPJ, part_NAME, "NÃO", cfop, vl_opr, vl_bc_icms, vl_icms, "2%", vl_ret]
                    outputsheet_rows.append(row)
    # Agora temos o valor total de RET e a planilha final.
    spreadsheet_header = ["Número", "Chave", "CNPJ", "Nome", "Contribuinte", "CFOP", "VL_OPR", "BC_ICMS", "VL_ICMS", "Alíquota RET", "ICMS RET"]
    workbook = xlsxwriter.Workbook(OUTPUT_FOLDER + 'Planilha de Apuração do RET.xlsx')
    worksheet = workbook.add_worksheet("ICMS RET")
    bold_header_format = workbook.add_format({'align':'center', 'bold':True, 'font_color':'white', 'bg_color':'black'})
    cell_format = workbook.add_format({'border':1})
    bold_format = workbook.add_format({'border':1, 'bold':True})
    worksheet.write_row("A1", spreadsheet_header, bold_header_format)
    i = 2
    for row in outputsheet_rows:
        worksheet.write_row("A"+str(i), row, cell_format)
        i += 1
    worksheet.write("J"+str(i), "TOTAL:", bold_header_format)
    worksheet.write("K"+str(i), VL_TOTAL_RET, bold_format)
    #
    column_widths = [12, 20, 16, 50, 12, 8, 10, 10, 10, 12, 10]
    for i in range(len(column_widths)):
        worksheet.set_column(i, i, column_widths[i])
    try:
        workbook.close()
    except:
        sg.popup("Não foi possível gerar a planilha. Você tem permissão para criá-la no local de saída? Ela está aberta?")
    # Feita a planilha, vamos agora gerar os registros da Sub-Apuração (registro 1900 e registros filhos)
    reg_1900 = ['1900', '3', 'Sub-Apuração e-PTA n. 45.000023237-88']
    reg_1910 = ['1910', efd_array[0][3], efd_array[0][4]]
    deb_str = set_value(VL_TOTAL_DEB)
    ret_str = set_value(VL_TOTAL_RET)
    diff_str = set_value(VL_TOTAL_DEB-VL_TOTAL_RET)
    reg_1920 = ['1920', deb_str, '0', '0', '0', diff_str, '0', '0', ret_str, '0', ret_str, '0', '0']
    reg_1921 = ['1921', 'MG020002', '', diff_str]
    reg_1926 = ['1926', '000', ret_str, '', '2196', '', '', '', '', efd_array[0][3][2:]]
    for i in range(len(efd_array)):
        if efd_array[i][0] == '1990':
            efd_array = efd_array[:i] + [reg_1900] + [reg_1910] + [reg_1920] + [reg_1921] + [reg_1926] + efd_array[i:]
            break
    efd_array = insert_9900_counter(efd_array, '1900')
    efd_array = insert_9900_counter(efd_array, '1910')
    efd_array = insert_9900_counter(efd_array, '1920')
    efd_array = insert_9900_counter(efd_array, '1921')
    efd_array = insert_9900_counter(efd_array, '1926')
    return efd_array

def remove_C170(efd_array):
    new_efd = []
    outgoing = False
    for row in efd_array:
        if row[0] == 'C100':
            outgoing = True if (row[2] == '0') else False
        if row[0] == 'C170' and outgoing:
            continue
        new_efd.append(row)
    return new_efd
    
def read_efd(efd_path):
    efd_array = []
    with open(efd_path, "r", encoding="latin-1") as efd:
        for line in efd:
            efd_array.append(line.replace("\n", "").split("|")[1:-1])
            if line.replace("\n", "").split("|")[1:-1][0] == "9999":
                break
    return efd_array

def write_efd(output_folder, efd_array):
    with open(output_folder + "EFD Saida.txt", "w+") as fp:
        new_efd = ""
        for line in efd_array:
            new_line = "|"
            for column in line:
                new_line += column + "|"
            new_line += "\n"
            new_efd += new_line
        fp.write(new_efd)

def main():
    # Defining the SimpleGUI Layout
    sg.theme("DarkGrey13")
    texts = [
        [sg.Text("Este programa fará o lançamento dos ajustes do Regime Especial de Tributação\nda empresa LAFEBER IND. E COM. DE CONDUTORES ELETRICOS LTDA.\n\n")],
        [sg.Text("EFD de entrada: ")],
        [sg.Text("Diretório onde deseja gerar a nova EFD: ")],
    ]
    header = [sg.Column([texts[0]], justification='c')]
    efd_filebrowser = [
        sg.In(size=(65, 1), enable_events=True, key="EFD"),
        sg.FileBrowse(button_text="Abrir"),
    ]
    output_folderbrowser = [
        sg.In(size=(60, 1), enable_events=True, key="OUTPUT_FOLDER"),
        sg.FolderBrowse(button_text="Selecionar"),
    ]
    progress_log = [sg.Output(size=(70, 5), key='-OUTPUT-')]

    send_button = [
        sg.Column([[sg.Button('Lançar Ajustes!', key="-SEND-")]], justification='c')]
    layout = [header, texts[1], efd_filebrowser,
              texts[2], output_folderbrowser, [sg.Text("\n")], progress_log, send_button]
    window = sg.Window("RET Lafeber", layout, grab_anywhere=False)
    while True:
        event, values = window.read()
        # End program if user closes window or
        # presses the OK button
        if event == "-SEND-":
            window.Element("-OUTPUT-").Update("")
            efd_path = values["EFD"]
            output_folder = values["OUTPUT_FOLDER"] + '/'
            efd_array = read_efd(efd_path)
            efd_array = ajustes_RET(efd_array)
            efd_array = subA_RET(efd_array, output_folder)
            efd_array = remove_C170(efd_array)
            efd_array = fix_unusedItems(efd_array)
            efd_array = update_counters(efd_array)
            write_efd(output_folder, efd_array)
            sg.popup("Fim do processo :)", title="Atenção")
        if event == sg.WIN_CLOSED:
            break
    window.close()
    return

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        sg.popup("ERRO NA EXECUÇÃO:", e, title="Erro!")
