import os
import time
import struct
import subprocess
import threading
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json

# xlwings é opcional: se não estiver instalado, o programa continua
# funcionando normalmente, só sem a sincronização com o Excel.
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False

# keyboard é opcional: usado só pro atalho global de sincronização (funciona
# mesmo com o programa em segundo plano/minimizado). Se não estiver
# instalado, o programa continua funcionando normalmente, só sem o atalho.
# Instalar com: pip install keyboard --break-system-packages
try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    KEYBOARD_AVAILABLE = False

# ==========================================
# BANCO DE DADOS: IDS E CURVAS DE EXP
# ==========================================
CURVA_1_IDS = {2, 9, 20, 42, 50, 52, 53, 55, 56, 81, 90, 96, 97, 103, 111, 112, 114, 136, 143, 151, 190, 208, 219, 231, 234, 241, 242, 303, 307, 316, 317, 320, 321, 322, 325, 343, 348, 361, 387, 388, 389, 391, 392, 437, 438, 457, 458, 466, 510, 512, 514, 515, 564, 567, 569, 573, 577, 578, 582, 594, 595, 607, 609, 611, 622, 623, 626, 629, 631, 682, 687, 697, 701, 705, 706, 707, 708, 709, 713, 728, 750, 790}
CURVA_2_IDS = {712, 714, 630, 91, 760, 365, 711, 15, 454, 590, 615, 220, 17, 314, 43, 752, 586, 579, 580, 13, 70, 375, 592, 228, 93, 598, 313, 30, 452, 5, 113, 209, 729, 399, 367, 78, 455, 341, 240, 394, 393, 100, 730, 3, 369, 10, 463, 115, 396, 373, 349, 102, 363, 474, 759, 621, 14, 368, 130, 347, 58, 372, 398, 395, 610, 51, 68, 601, 377, 758, 77, 72, 456, 25, 548, 370, 371, 702, 574, 397, 137, 308, 222, 92, 755, 698, 597, 16, 54, 218, 326, 235, 191, 12, 710, 304, 87, 11, 344, 183, 139, 141, 146, 716, 147, 144, 142, 145, 162, 680, 676, 681, 163, 679, 670, 669, 390, 715, 327, 627, 408, 23, 720, 723, 133, 726, 107, 364, 134, 301, 602, 226, 31, 411, 177, 753, 584, 576, 376, 770, 129, 217, 908, 6, 404, 379, 76, 907, 211, 727, 402, 380, 149, 401, 311, 85, 199, 237, 198, 359, 73, 413, 132, 412, 74, 465, 59, 724, 82, 410, 28, 596, 41, 606, 718, 407, 33, 378, 406, 374, 405, 232, 116, 239, 79, 4, 44, 229, 683, 756, 731, 920, 605, 140, 479, 722, 71, 721, 61, 302, 588, 101, 21, 138, 575, 192, 409, 414, 699, 84, 345, 484, 719, 224, 223, 491, 26, 403, 309, 148, 1, 400, 210, 342, 305, 186, 187, 185, 188, 184, 485, 487, 486, 488}
CURVA_3_IDS = {182, 175, 358, 468, 432, 418, 689, 449, 690, 632, 423, 773, 346, 419, 732, 703, 34, 37, 383, 171, 170, 80, 312, 47, 385, 230, 747, 616, 439, 35, 195, 233, 214, 180, 69, 196, 227, 428, 427, 36, 767, 179, 126, 38, 128, 440, 745, 771, 912, 225, 46, 32, 127, 194, 492, 422, 150, 744, 338, 600, 451, 617, 742, 57, 603, 434, 135, 95, 431, 450, 429, 60, 19, 743, 739, 738, 733, 98, 749, 425, 424, 189, 176, 83, 704, 754, 688, 178, 725, 740, 168, 174, 24, 748, 453, 310, 306, 49, 426, 75, 700, 741, 315, 213, 86, 614, 461, 737, 94, 48, 735, 774, 45, 197, 734, 613, 421, 193, 27, 172, 117, 417, 416, 173, 382, 751, 675, 678, 677, 169, 448, 692, 691, 693}
CURVA_4_IDS = {357, 841, 328, 840, 106, 435, 181, 40, 775, 105, 915, 779, 777, 215, 778, 104, 776, 604, 118, 772, 757, 88, 420, 494, 766, 39}

# Experiência base (Nível 1 ao 99)
EXP_CURVA_1 = [0, 8, 88, 278, 666, 1287, 1995, 2919, 4086, 5526, 7269, 9345, 11781, 14604, 17844, 21531, 25695, 30363, 35562, 41322, 47673, 54645, 62265, 70560, 79560, 89295, 99795, 111087, 123198, 136158, 149997, 164745, 180429, 197076, 214716, 233379, 253095, 273891, 295794, 318834, 343041, 368448, 395082, 422967, 452133, 482610, 514428, 547614, 582195, 618201, 655662, 694608, 735066, 777063, 820629, 865794, 912588, 961038, 1011171, 1063017, 1116606, 1171968, 1229130, 1288119, 1348965, 1411698, 1476348, 1542942, 1611507, 1682073, 1754670, 1829328, 1906614, 1984935, 2065941, 2149122, 2234508, 2322126, 2412003, 2504169, 2598654, 2695488, 2794698, 2896311, 3000357, 3106866, 3215868, 3327390, 3441459, 3558105, 3677358, 3799248, 3923802, 4051047, 4181013, 4313730, 4449228, 4587534, 4800000]
EXP_CURVA_2 = [0, 8, 130, 375, 845, 1629, 2517, 3675, 5139, 6945, 9129, 11727, 14775, 18309, 22365, 26979, 32187, 38025, 44529, 51735, 59679, 68397, 77925, 88299, 99555, 111729, 124857, 138975, 154119, 170325, 187629, 206067, 225675, 246489, 268545, 291879, 316527, 342525, 369909, 398715, 428979, 460737, 494025, 528879, 565335, 603429, 643197, 684675, 727899, 772905, 819729, 868407, 918975, 971469, 1025925, 1082379, 1140867, 1201425, 1264089, 1328895, 1395879, 1465077, 1536525, 1610259, 1686315, 1764729, 1845537, 1928775, 2014479, 2102685, 2193429, 2286747, 2382675, 2481249, 2582505, 2686479, 2793207, 2902725, 3015069, 3130275, 3248379, 3369417, 3493425, 3620439, 3750495, 3883629, 4019877, 4159275, 4301859, 4447665, 4596729, 4749087, 4904775, 5063829, 5226285, 5392179, 5561547, 5734425, 6000000]
EXP_CURVA_3 = [0, 8, 158, 439, 962, 1860, 2880, 4209, 5889, 7962, 10470, 13455, 16959, 21021, 25683, 30987, 36972, 43683, 51162, 59448, 68583, 78606, 89559, 101487, 114429, 128427, 143523, 159756, 177168, 195801, 215697, 236898, 259446, 283380, 308742, 335574, 363915, 393810, 425301, 458427, 493230, 529749, 568026, 608106, 650028, 693834, 739566, 787263, 836967, 888720, 942564, 998541, 1056693, 1117059, 1179681, 1244601, 1311858, 1381497, 1453560, 1528089, 1605126, 1684707, 1766874, 1851672, 1939140, 2029320, 2122254, 2217981, 2316543, 2417982, 2522340, 2629659, 2739981, 2853345, 2969793, 3089367, 3212106, 3338055, 3467256, 3599748, 3735573, 3874770, 4017381, 4163451, 4313019, 4466127, 4622817, 4783128, 4947102, 5114781, 5286207, 5461422, 5640468, 5823384, 6010212, 6200994, 6395769, 6594582, 6900000]
EXP_CURVA_4 = [0, 8, 184, 492, 1048, 2022, 3129, 4575, 6402, 8655, 11382, 14628, 18435, 22848, 27915, 33681, 40188, 47481, 55608, 64614, 74541, 85434, 97341, 110307, 124374, 139587, 155994, 173640, 192567, 212820, 234447, 257493, 282000, 308013, 335580, 364746, 395553, 428046, 462273, 498279, 536106, 575799, 617406, 660972, 706539, 754152, 803859, 855705, 909732, 965985, 1024515, 1085367, 1148580, 1214199, 1282272, 1352844, 1425957, 1501656, 1579989, 1661001, 1744734, 1831233, 1920546, 2012718, 2107791, 2205810, 2306823, 2410875, 2518008, 2628267, 2741700, 2858352, 2978265, 3101484, 3228057, 3358029, 3491442, 3628341, 3768774, 3912786, 4060419, 4211718, 4366731, 4525503, 4688079, 4854498, 5024808, 5199060, 5377293, 5559552, 5745885, 5936337, 6130950, 6329769, 6532842, 6740214, 6951927, 7168026, 7500000]

def get_exp_needed(digimon_id, level):
    """Retorna a experiência total necessária para atingir um dado level."""
    if level < 1: level = 1
    if level > 99: level = 99
    idx = level - 1
    if digimon_id in CURVA_1_IDS: return EXP_CURVA_1[idx]
    if digimon_id in CURVA_2_IDS: return EXP_CURVA_2[idx]
    if digimon_id in CURVA_3_IDS: return EXP_CURVA_3[idx]
    if digimon_id in CURVA_4_IDS: return EXP_CURVA_4[idx]
    return 0 

# ==========================================
# CONFIGURAÇÃO FIXA
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(SCRIPT_DIR)

SAVE_FILE_DEC = "save_dec.bin"
CONFIG_FILE = "config.json"
AES_KEY = "33393632373736373534353535383833"
DIGIMON_SIZE = 336

# ==========================================
# INTEGRAÇÃO COM EXCEL (xlwings)
# ==========================================
EXCEL_SYNC_ENABLED = False
COMPARADOR_SYNC = False  # Se True, mostra o botão "📊 Sync Comparator" (aba "Comparações_Talento"). Se False, o botão nem aparece.
AUTO_SYNC_TALENTOS_VISIBLE = False  # Se False, esconde o auto-sync e impede qualquer sync automatico com Excel.
AUTO_SYNC_TALENTOS_DEFAULT = False # Valor inicial quando ainda nao existe preferencia salva no config.json.


# Caminho genérico de exemplo
EXCEL_FILE_PATH = r"C:\caminho\para\sua\planilha.xlsm" 

try:
    from meus_segredos import MEU_CAMINHO_EXCEL
    EXCEL_FILE_PATH = MEU_CAMINHO_EXCEL
except ImportError:
    pass  # Se o arquivo não existir, ignora e usa o caminho genérico acima

EXCEL_SHEET_NAME = "Digimons"  # Nome FIXO da aba principal (Coluna F/V/W). Ajuste aqui se o nome da sua aba for outro.
# IMPORTANTE: antes esse valor era None, o que fazia o código usar
# book.sheets.active (a aba que estivesse VISÍVEL no Excel no momento do
# clique). Isso causava dois problemas: (1) se você estivesse numa aba
# diferente (ex.: "Comparações_Talento"), o sync procurava os nomes na
# Coluna F *dessa* aba errada, não achava nada e voltava sem fazer nada -
# parecendo "rápido" mas na real não sincronizando. (2) escrever numa aba
# que está ativa/visível é mais lento (Excel atualiza seleção/formatação
# condicional a cada escrita) do que escrever numa aba em segundo plano.
# Fixando o nome, o sync sempre mira a aba certa, esteja ela visível ou não.
EXCEL_COL_NAME = 6           # Coluna F (nome do Digimon)
EXCEL_COL_ASCENDANT = 22     # Coluna V (ascendant_talent_raw)
EXCEL_COL_LEVEL = 23         # Coluna W (level) 
EXCEL_COL_ELO = 24           # Coluna X (EloS)
EXCEL_HEADER_ROW = 1         # Primeira linha considerada dado (pule se tiver cabeçalho maior)

# Aba auxiliar que guarda o histórico de talento ascendente ao longo do tempo,
# pra ajudar a entender como esse status se comporta.
EXCEL_COMPARISON_SHEET_NAME = "Comparações_Talento"
MAX_COMPARACOES = 4          # Total de comparações na tabela de comparações de talento

# ==========================================
# ATALHO GLOBAL DE TECLADO (opcional)
# ==========================================
# Dispara o botão azul "🔄 Sync Main Sheet" com uma tecla, mesmo com o
# programa em segundo plano/minimizado (não precisa a janela estar em foco).
# Isso usa a lib "keyboard" (hook de teclado em baixo nível do Windows), que
# é opcional - se não estiver instalada, o atalho simplesmente não é
# registrado e o resto do programa funciona normal.
#
# Recomendo usar F13-F20: são teclas que praticamente não existem em teclado
# físico nenhum, então nunca vão conflitar com outro atalho seu. Dá pra
# mapear pra um botão de controle/joystick usando um software de macro (ex.:
# reWASD, AutoHotkey, JoyToKey, etc.) apontando pra uma dessas teclas.
#
# Nomes aceitos pela lib keyboard: "f13", "f14", ..., "f24". Combinações
# também funcionam, ex.: "ctrl+f13", "ctrl+shift+f13".
HOTKEY_SYNC_ENABLED = True
HOTKEY_FOCUS_WISHLIST = "f19"  # <- foca direto na busca da Wishlist
HOTKEY_SYNC_TALENTOS = "f18"   # <- troque aqui pra tecla que você quiser

# CORES DARK MODE
BG_COLOR = "#121212"
FG_COLOR = "#00FF00"  
FG_ALERT = "#FF4500"  
FG_ALMOST = "#FFD700" 
PANEL_BG = "#1E1E1E"
BTN_BG = "#333333"

# Ciclo de limites para a lista "Quase lá": altere apenas os dois primeiros valores se quiser mudar o ciclo.
QUASE_LIST_CYCLE = [12, 32, 0]  # 0 = mostrar todos

# A faixa fixa da barrinha vai até quanto de xp? Altere aqui.
MAX_LEVEL_BARRA = 35000

# A faixa de espaços para o truncamento inteligente de nomes para as barras ficarem alinhadas:
MAX_NAME_LEN = 18

# Numero de valores para a lista flutuante na textbox de busca da Wishlist
VALORES_BUSCA = 40

# ==========================================
# DICIONÁRIO DE IDIOMAS (i18n)
# ==========================================
I18N = {
    "EN": {
        "settings": "⚙️ SETTINGS",
        "atingiu": "<-- REACHED THE CAP!",
        "cap_list_title": " 🏁 REACHED THE CAP:",
        "quase": " ⏳ ALMOST THERE. (Closest to reach the limit)",
        "modo_auto": "Automatic Mode",
        "modo_manual": "Manual Mode (Ignore auto-save)",
        "btn_change_folder": "📂 Change Save Folder",
        "btn_sync_talentos": "🔄 Sync Main Sheet",
        "btn_sync_talentos_running": "🔄 Syncing...",
        "btn_sync_comparador": "📊 Sync Comparator",
        "btn_sync_comparador_running": "📊 Syncing...",
        "auto_sync_on": "Auto Sync ON",
        "auto_sync_off": "Auto Sync OFF",
        "msg_auto_sync_on": "✅ Auto-sync enabled. It will run after each new save.",
        "msg_auto_sync_off": "ℹ️ Auto-sync disabled.",
        "msg_sync_excel_ok": "✅ Excel updated: ",
        "msg_sync_excel_none": "ℹ️ No protected Digimon matched in the spreadsheet.",
        "msg_sync_excel_disabled": "⚠️ Excel sync is disabled (EXCEL_SYNC_ENABLED = False).",
        "msg_sync_excel_no_data": "⚠️ No save data loaded yet. Read a save first.",
        "msg_hotkey_registered": "⌨️ Global hotkey active: {key} (works in background).",
        "msg_hotkey_failed": "⚠️ Failed to register global hotkey ({key}): ",
        "msg_hotkey_no_lib": "ℹ️ \"keyboard\" lib not installed - global hotkey disabled (pip install keyboard).",
        "msg_comparador_disabled": "⚠️ Comparator is disabled\n(COMPARADOR_SYNC = False).",
        "msg_comparison_reset": " 🔄 Cycle reset! Last talents moved to columns A & B ({count}).\nNext click will start new comparisons.",
        "msg_comparison_created": " 📊 \"Comparações_Talento\" tab created\nand populated ({count}).",
        "msg_comparison_updated": " 📊 Comparison round added to\n\"Comparações_Talento\" ({count} updated, {new} new).",
        "msg_comparison_error": " ⚠️ [Comparison Sheet]\nFailed to update: ",
        "lbl_inspect": "Inspect Save:",
        "lbl_paused": "⚠️ PAUSED",
        "btn_resume": "RESUME TRACKING",
        "title_initial_config": "Initial Configuration",
        "msg_select_folder": "Please select the folder where the Digimon Story saves are located.",
        "dialog_select_folder_title": "Select the Digimon saves folder",
        "title_exit": "Exit",
        "msg_need_folder": "The program needs the save folder to work. Do you want to exit?",
        "dialog_change_folder_title": "Select the new saves folder",
        "msg_folder_changed": "Save folder changed to:",
        "err_critical_title": " ❌ CRITICAL ERROR: MISSING DEPENDENCY",
        "err_openssl_notfound": "\n The file 'openssl.exe' was not found.\n",
        "err_openssl_path": " It needs to be EXACTLY in this folder:\n",
        "err_decrypt": "\n ❌ ERROR while trying to decrypt the save:\n",
        "app_title": "      🎮 DIGIMON LEVEL CAP CHECKER - REAL-TIME UI",
        "status_inspecting": "Inspecting Save: ",
        "status_monitoring": "Monitoring latest Save: ",
        "status_prefix": " 🟢 Status: ",
        "lbl_updated": " | Updated: ",
        "lbl_summary": " ⚠️  Summary:",
        "msg_all_normal": "    All Digimons are evolving normally.",
        "summary_alerts": " digimon(s) reached the Maximum limit",
        "summary_almost": " digimon(s) need 1 more level to reach the Max",
        "summary_lv99": " digimon(s) are at level 99",
        "summary_protected": " digimon(s) are protected (locked)",
        "btn_hide_details": "Hide details",
        "btn_show_lv99": "Click to see the Level 99 digimons",
        "btn_show_almost": "Click to see who's almost there",
        "btn_show_remaining": "View remaining digimons",
        "btn_hide_remaining": "Hide remaining digimons",
        "remaining_list_title": " 🔍 REMAINING DIGIMONS:",
        "btn_show_wishlist": "Click to see the Wishlist",
        "btn_quase_show_until_10": f"Show up to {QUASE_LIST_CYCLE[0]} Digimons",
        "btn_quase_show_until_30": f"Show up to {QUASE_LIST_CYCLE[1]} Digimons",
        "btn_quase_show_until_all": "Show all digimons",
        "btn_quase_show_until_custom": "Show up to {count} Digimons",
        "msg_quase_invalid_count": "Enter a valid number for the amount.",
        "lbl_quase_or_show_until": "or Show up to:",
        "lbl_quase_search": "Search Digimon:",
        "btn_quase_fetch": "🔍 Search",
        "radio_all": "All",
        "radio_party": "Party",
        "radio_box": "Box",
        "radio_farm": "Farm",
        "checkbox_protected": "Protected",
        "checkbox_wishlist": "Wishlist",
        "checkbox_wishlist_help": "Include wishlist targets in the 'Almost there' list",
        "msg_quase_no_results": "No digimons match the applied filters.",
        "lbl_lv99_title": " 👑 DIGIMONS AT MAXIMUM LEVEL (99):",
        "waiting_msg": "\n[Waiting for game update... Keep this open on the 2nd screen]",
        "paused_msg": "\n[TRACKING PAUSED. Click the red side button to return to the radar]",
        "lvl_abbr": "Lv.",
        "limite_abbr": "Cap",
        "faltam_abbr": "Missing:",
        "wishlist_title": " 🎯 TARGET WISHLIST / EVOLUTION GOALS:",
        "target_reached": "<-- TARGET REACHED!",
        "wishlist_auto_removed": "⚠️  WISHLIST: digimon(s) not found in the current save (evolved / released / different save):",
        "wishlist_panel_title": "🎯 WISHLIST / TARGET TRACKER",
        "wishlist_search_label": "Search:",
        "wishlist_search_btn": "🔍 Search",
        "wishlist_choose_label": "Choose:",
        "wishlist_target_label": "Target Lv:",
        "wishlist_add_btn": "➕ Add Target",
        "msg_warning_title": "Warning",
        "msg_search_title": "Search",
        "wishlist_err_empty_query": "Please type an ID or part of a Name to search.",
        "wishlist_err_no_save": "No save has been loaded yet.",
        "wishlist_no_results": "No Digimon found for the term '{query}'.",
        "wishlist_err_select_first": "Select a Digimon from the search list first.",
        "wishlist_err_invalid_target": "Enter a valid Target Level (e.g. 60, 99).",
        "wishlist_err_target_range": "The target level must be between 1 and 99.",
        "wishlist_err_target_too_high": "The Target Level ({tgt}) cannot be GREATER than the Digimon's current Talent/Cap ({cap}).",
        "btn_wishlist_readd": "🔄 Re-add",
        "btn_wishlist_forget": "❌ Remove",
        "wishlist_not_found_msg": "'{name}' was not present in your inventory. The save was likely overwritten or deleted.",
        "wishlist_readded_msg": "'{name}' is present in the current save again and was restored to the wishlist.",
        "wishlist_ctx_paused": "You can add Digimon from this save ({save}) to the Wishlist below:",
        "target_abbr": "Target",
        "radio_sort_xp": "Least XP Remaining",
        "radio_sort_acc": "Lower talent value Acc",
        "lbl_obs_acc": "Note: The values ​​in braces {} on the lines represent the new talent value if you evolve the Digimon.",
        "loc_labels": {"PARTY": "PARTY", "BOX": "BOX", "FAZENDA": "FARM"}
    },
    "PT": {
        "settings": "⚙️ CONFIGURAÇÕES",
        "atingiu": "<-- ATINGIU O LIMITE!",
        "cap_list_title": " 🏁 CHEGOU NO CAP:",
        "quase": " ⏳ QUASE LÁ. (Mais próximos de chegar no limite)",
        "modo_auto": "Modo Automático",
        "modo_manual": "Modo Manual (Ignora o Auto-Save)",
        "btn_change_folder": "📂 Mudar Pasta de Saves",
        "btn_sync_talentos": "🔄 Sync Planilha Principal",
        "btn_sync_talentos_running": "🔄 Sincronizando...",
        "btn_sync_comparador": "📊 Sync Comparador",
        "btn_sync_comparador_running": "📊 Sincronizando...",
        "auto_sync_on": "Auto Sync ON",
        "auto_sync_off": "Auto Sync OFF",
        "msg_auto_sync_on": "✅ Auto-sync ativado. Vai rodar a cada novo save.",
        "msg_auto_sync_off": "ℹ️ Auto-sync desativado.",
        "msg_sync_excel_ok": "✅ Excel atualizado: ",
        "msg_sync_excel_none": "ℹ️ Nenhum Digimon protegido encontrado na planilha.",
        "msg_sync_excel_disabled": "⚠️ Sync com Excel está desativado (EXCEL_SYNC_ENABLED = False).",
        "msg_sync_excel_no_data": "⚠️ Nenhum save carregado ainda. Leia um save primeiro.",
        "msg_hotkey_registered": "⌨️ Atalho global ativo: {key} (funciona em segundo plano).",
        "msg_hotkey_failed": "⚠️ Falha ao registrar o atalho global ({key}): ",
        "msg_hotkey_no_lib": "ℹ️ Lib \"keyboard\" não instalada - atalho global desativado (pip install keyboard).",
        "msg_comparador_disabled": "⚠️ Comparador está desativado\n(COMPARADOR_SYNC = False).",
        "msg_comparison_reset": " 🔄 Ciclo reiniciado! Últimos talentos movidos para as colunas A e B ({count}).\nO próximo clique iniciará novas comparações.",
        "msg_comparison_created": " 📊 Aba \"Comparações_Talento\" criada\ne populada ({count}).",
        "msg_comparison_updated": " 📊 Rodada de comparação adicionada em\n\"Comparações_Talento\" ({count} atualizados, {new} novos).",
        "msg_comparison_error": " ⚠️ [Aba de Comparação]\nFalha ao atualizar: ",
        "lbl_inspect": "Inspecionar Save:",
        "lbl_paused": "⚠️ PAUSADO",
        "btn_resume": "VOLTAR A RASTREAR",
        "title_initial_config": "Configuração Inicial",
        "msg_select_folder": "Por favor, selecione a pasta onde os saves de Digimon Story estão localizados.",
        "dialog_select_folder_title": "Selecione a pasta dos saves do Digimon",
        "title_exit": "Sair",
        "msg_need_folder": "O programa precisa da pasta de saves para funcionar. Deseja sair?",
        "dialog_change_folder_title": "Selecione a nova pasta de saves",
        "msg_folder_changed": "Pasta de saves alterada para:",
        "err_critical_title": " ❌ ERRO CRÍTICO: DEPENDÊNCIA AUSENTE",
        "err_openssl_notfound": "\n O arquivo 'openssl.exe' não foi encontrado.\n",
        "err_openssl_path": " Ele precisa estar EXATAMENTE nesta pasta:\n",
        "err_decrypt": "\n ❌ ERRO ao tentar descriptografar o save:\n",
        "app_title": "      🎮 DIGIMON LEVEL CAP CHECKER - UI EM TEMPO REAL",
        "status_inspecting": "Inspecionando Save: ",
        "status_monitoring": "Monitorando Save mais recente: ",
        "status_prefix": " 🟢 Status: ",
        "lbl_updated": " | Atualizado: ",
        "lbl_summary": " ⚠️  Resumo:",
        "msg_all_normal": "    Todos os Digimons estao evoluindo normalmente.",
        "summary_alerts": " digimon(s) atingiram o limite Máximo",
        "summary_almost": " digimon(s) falta 1 level para o limite Máximo",
        "summary_lv99": " digimon(s) estão no level 99",
        "summary_protected": " digimon(s) estão protegidos (cadeado)",
        "btn_hide_details": "Ocultar detalhes",
        "btn_show_lv99": "Clique para ver os digimons no Nv. 99",
        "btn_show_almost": "Clique para ver quem está quase lá",
        "btn_show_remaining": "Visualizar o restante dos digimons",
        "btn_hide_remaining": "Ocultar Digimons restantes",
        "remaining_list_title": " 🔍 RESTANTE DOS DIGIMONS:",
        "btn_show_wishlist": "Clique para ver a Wishlist",
        "lbl_lv99_title": " 👑 DIGIMONS NO LEVEL MÁXIMO (99):",
        "waiting_msg": "\n[Aguardando atualização do jogo... Mantenha aberto na 2ª tela]",
        "paused_msg": "\n[RASTREAMENTO PAUSADO. Clique no botão lateral vermelho para voltar ao radar]",
        "lvl_abbr": "Nv.",
        "limite_abbr": "Limite",
        "faltam_abbr": "Faltam:",
        "wishlist_title": " 🎯 WISHLIST / METAS DE EVOLUÇÃO:",
        "target_reached": "<-- META ATINGIDA!",
        "wishlist_auto_removed": "⚠️  WISHLIST: digimon(s) não encontrado(s) no save atual (evoluiu / foi liberado / save diferente):",
        "wishlist_panel_title": "🎯 WISHLIST / META TRACKER",
        "wishlist_search_label": "Busca:",
        "wishlist_search_btn": "🔍 Buscar",
        "wishlist_choose_label": "Escolha:",
        "wishlist_target_label": "Nível Alvo:",
        "wishlist_add_btn": "➕ Add Meta",
        "msg_warning_title": "Aviso",
        "msg_search_title": "Pesquisa",
        "wishlist_err_empty_query": "Por favor, digite uma ID ou parte do Nome para pesquisar.",
        "wishlist_err_no_save": "Nenhum save foi carregado ainda.",
        "wishlist_no_results": "Nenhum Digimon encontrado para o termo '{query}'.",
        "wishlist_err_select_first": "Selecione um Digimon na lista de pesquisa primeiro.",
        "wishlist_err_invalid_target": "Digite um Nível Alvo válido (ex: 60, 99).",
        "wishlist_err_target_range": "O nível alvo deve estar entre 1 e 99.",
        "wishlist_err_target_too_high": "O Target Level ({tgt}) não pode ser MAIOR que o Talento/Limite atual do Digimon ({cap}).",
        "btn_wishlist_readd": "🔄 Readicionar",
        "btn_wishlist_forget": "❌ Remover",
        "wishlist_not_found_msg": "'{name}' não está presente no seu inventário. Provavelmente o save foi sobrescrito ou deletado.",
        "wishlist_readded_msg": "'{name}' está presente no save atual novamente e foi restaurado na wishlist.",
        "wishlist_ctx_paused": "Você pode adicionar Digimons deste save ({save}) na Wishlist aqui embaixo:",
        "btn_quase_show_until_10": f"Mostrar até {QUASE_LIST_CYCLE[0]} Digimons",
        "btn_quase_show_until_30": f"Mostrar até {QUASE_LIST_CYCLE[1]} Digimons",
        "btn_quase_show_until_all": "Mostrar todos os Digimons",
        "btn_quase_show_until_custom": "Mostrar até {count} Digimons",
        "msg_quase_invalid_count": "Digite um número válido para a quantidade.",
        "lbl_quase_or_show_until": "ou Mostrar até:",
        "lbl_quase_search": "Pesquisar Digimon:",
        "btn_quase_fetch": "🔍 Buscar",
        "radio_all": "Todos",
        "radio_party": "Party",
        "radio_box": "Box",
        "radio_farm": "Fazenda",
        "checkbox_protected": "Protegidos",
        "checkbox_wishlist": "Wishlist",
        "checkbox_wishlist_help": "Inclui os Digimon da Wishlist na lista atual",
        "msg_quase_no_results": "Nenhum digimon encontrado com a busca.",
        "target_abbr": "Alvo",
        "radio_sort_xp": "Menor XP Restante",
        "radio_sort_acc": "Menor Talento ACC",
        "lbl_obs_acc": "OBS: Os valores entre chaves{} nas linhas representam o novo valor de talento se você evoluir o Digimon.",
        "loc_labels": {"PARTY": "PARTY", "BOX": "BOX", "FAZENDA": "FAZENDA"}
    }
}

class DigimonMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Digimon Level Cap Checker")
        self.root.configure(bg=BG_COLOR)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.config_data = {}
        self.save_dir = ""
        self.wishlist = []
        self.search_results_map = []
        
        if not self.initialize_config():
            return
            
        self.mode = tk.StringVar(value="AUTO")
        self.is_paused = False
        self.last_mtime = 0
        self._pending_save_check = None  # (filepath, mtime, size) aguardando confirmação de "arquivo parou de mudar"
        self.blink_state = False
        self.auto_sync_talentos_enabled = (
            AUTO_SYNC_TALENTOS_VISIBLE
            and self.config_data.get("auto_sync_talentos_enabled", AUTO_SYNC_TALENTOS_DEFAULT)
        )
        self.auto_sync_talentos_var = tk.StringVar(value="ON" if self.auto_sync_talentos_enabled else "OFF")
        self._last_auto_sync_signature = None
        # ... dentro do __init__, junto das outras variáveis ...
        self.quase_sort_by = self.config_data.get("quase_sort_by", "XP") # XP ou ACC

        self.setup_ui()
        self.setup_global_hotkey()
        self.update_loop()

    def initialize_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    self.config_data = json.load(f)
            except Exception:
                self.config_data = {}

        self.lang = self.config_data.get("language", "PT")
        self.lang_var = tk.StringVar(value=self.lang)
        self.wishlist = self.config_data.get("wishlist", [])

        # Estado (aberto/fechado) de cada lista expansível, lembrado entre sessões
        self.show_almost = self.config_data.get("show_almost", False)
        self.show_lv99 = self.config_data.get("show_lv99", False)
        self.show_wishlist = self.config_data.get("show_wishlist", False)
        self.show_remaining = False
        self.quase_list_max_count = self.config_data.get("quase_list_max_count", QUASE_LIST_CYCLE[0])
        self.quase_filter_location = self.config_data.get("quase_filter_location", "TODOS")
        self.quase_filter_protected = self.config_data.get("quase_filter_protected", False)
        self.quase_filter_wishlist = self.config_data.get("quase_filter_wishlist", False)

        # Ordem das listas reordenáveis (2, 3, 4). A 1ª lista (Level Cap) é sempre fixa.
        self.list_order = self.config_data.get("list_order", ["almost", "lv99", "wishlist"])

        if "geometry" in self.config_data:
            self.root.geometry(self.config_data["geometry"])
        else:
            self.root.geometry("1000x650")

        self.save_dir = self.config_data.get("save_dir", "")
        
        while not self.save_dir or not os.path.exists(self.save_dir):
            self.root.withdraw() 
            t = I18N[self.lang]
            messagebox.showinfo(t["title_initial_config"], t["msg_select_folder"])
            escolha = filedialog.askdirectory(title=t["dialog_select_folder_title"])
            
            if not escolha:
                if messagebox.askyesno(t["title_exit"], t["msg_need_folder"]):
                    self.root.destroy()
                    return False
            else:
                self.save_dir = os.path.normpath(escolha)
                self.config_data["save_dir"] = self.save_dir
                self.save_config()
                
        self.root.deiconify() 
        return True

    def save_config(self):
        self.config_data["geometry"] = self.root.geometry()
        self.config_data["save_dir"] = self.save_dir
        self.config_data["language"] = self.lang
        self.config_data["wishlist"] = self.wishlist
        self.config_data["show_almost"] = getattr(self, 'show_almost', False)
        self.config_data["show_lv99"] = getattr(self, 'show_lv99', False)
        self.config_data["show_wishlist"] = getattr(self, 'show_wishlist', False)
        self.config_data["quase_list_max_count"] = getattr(self, 'quase_list_max_count', QUASE_LIST_CYCLE[0])
        self.config_data["quase_filter_location"] = getattr(self, 'quase_filter_location', "TODOS")
        self.config_data["quase_filter_protected"] = getattr(self, 'quase_filter_protected', False)
        self.config_data["quase_filter_wishlist"] = getattr(self, 'quase_filter_wishlist', False)
        self.config_data["list_order"] = self.get_normalized_list_order()
        self.config_data["auto_sync_talentos_enabled"] = (
            AUTO_SYNC_TALENTOS_VISIBLE and getattr(self, 'auto_sync_talentos_enabled', False)
        )
        # ... dentro do save_config ...
        self.config_data["quase_sort_by"] = getattr(self, 'quase_sort_by', "XP")
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config_data, f, indent=4)
        except Exception:
            pass

    def toggle_remaining(self, filepath, filename):
        self.show_remaining = not self.show_remaining
        self.process_save(filepath, filename)

    def on_closing(self):
        self.save_config()
        self.teardown_global_hotkey()
        self.root.destroy()

    def get_normalized_list_order(self):
        """Garante que self.list_order sempre contenha exatamente as 3 chaves válidas
        (almost, lv99, wishlist), preservando a ordem já salva, removendo qualquer chave
        antiga que não exista mais (ex: 'protected', de uma versão anterior do programa),
        e anexando no final qualquer chave nova que ainda não existisse."""
        valid_keys = ["almost", "lv99", "wishlist"]
        current = getattr(self, 'list_order', None) or []
        normalized = [k for k in current if k in valid_keys]
        for k in valid_keys:
            if k not in normalized:
                normalized.append(k)
        self.list_order = normalized
        return normalized

    def move_list_order(self, key, direction):
        """Move uma lista reordenável (2ª a 4ª) uma posição pra cima (-1) ou pra baixo (+1)."""
        order = self.get_normalized_list_order()
        if key not in order:
            return
        idx = order.index(key)
        new_idx = idx + direction
        if 0 <= new_idx < len(order):
            order[idx], order[new_idx] = order[new_idx], order[idx]
            self.list_order = order
            self.save_config()
            if hasattr(self, '_current_filepath') and os.path.exists(self._current_filepath):
                self.process_save(self._current_filepath, self._current_filename)

    def change_directory(self):
        t = I18N[self.lang]
        nova_pasta = filedialog.askdirectory(title=t["dialog_change_folder_title"], initialdir=self.save_dir)
        if nova_pasta:
            self.save_dir = os.path.normpath(nova_pasta)
            self.save_config()
            self.last_mtime = 0
            self._pending_save_check = None
            self.update_combo_list()
            self.log(f"{t['msg_folder_changed']}\n{self.save_dir}", "header")
            self.root.focus_set()

    def on_lang_change(self):
        novo_idioma = self.lang_var.get()
        if novo_idioma != self.lang:
            self.lang = novo_idioma
            self.save_config()
            self.apply_language()

    def apply_language(self):
        """Atualiza o texto de todos os widgets estáticos para o idioma atual."""
        t = I18N[self.lang]
        self.lbl_settings_title.config(text=t["settings"])
        self.rb_auto.config(text=t["modo_auto"])
        self.rb_manual.config(text=t["modo_manual"])
        self.btn_change_folder.config(text=t["btn_change_folder"])
        if not getattr(self, "_sync_talentos_running", False):
            self.btn_sync_talentos.config(text=t["btn_sync_talentos"])
        if not getattr(self, "_sync_comparador_running", False):
            self.btn_sync_comparador.config(text=t["btn_sync_comparador"])
        if hasattr(self, "rb_auto_sync_on"):
            self.rb_auto_sync_on.config(text=t["auto_sync_on"])
            self.rb_auto_sync_off.config(text=t["auto_sync_off"])
        self.lbl_inspect.config(text=t["lbl_inspect"])
        self.lbl_paused.config(text=t["lbl_paused"])
        self.btn_resume.config(text=t["btn_resume"])
        self.lbl_wishlist_title.config(text=t["wishlist_panel_title"])
        self.lbl_wishlist_search.config(text=t["wishlist_search_label"])
        self.lbl_wishlist_choose.config(text=t["wishlist_choose_label"])
        self.btn_wish_search.config(text=t["wishlist_search_btn"])
        self.lbl_wishlist_target.config(text=t["wishlist_target_label"])
        self.btn_wish_add.config(text=t["wishlist_add_btn"])
        self.lbl_summary_title.config(text=t["lbl_summary"])
        if hasattr(self, "chk_quase_protected"):
            self.chk_quase_protected.config(text=t["checkbox_protected"])
        if hasattr(self, "chk_quase_wishlist"):
            self.chk_quase_wishlist.config(text=t["checkbox_wishlist"])
            self._quase_wishlist_hint_text = t["checkbox_wishlist_help"]

        if hasattr(self, "_current_filepath") and os.path.exists(self._current_filepath):
            self.process_save(self._current_filepath, self._current_filename)

    def _show_quase_wishlist_hint(self, event, text=None):
        if not hasattr(self, "_quase_wishlist_tooltip"):
            self._quase_wishlist_tooltip = tk.Toplevel(self.root)
            self._quase_wishlist_tooltip.withdraw()
            self._quase_wishlist_tooltip.overrideredirect(True)
            self._quase_wishlist_tooltip_label = tk.Label(
                self._quase_wishlist_tooltip,
                text="",
                bg="#333333",
                fg="white",
                padx=8,
                pady=4,
                relief="solid",
                borderwidth=1,
                font=("Consolas", 8),
            )
            self._quase_wishlist_tooltip_label.pack()
        if text is None:
            text = getattr(self, "_quase_wishlist_hint_text", "")
        self._quase_wishlist_tooltip_label.config(text=text)
        self._quase_wishlist_tooltip.geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
        self._quase_wishlist_tooltip.deiconify()

    def _hide_quase_wishlist_hint(self, event=None):
        if hasattr(self, "_quase_wishlist_tooltip"):
            self._quase_wishlist_tooltip.withdraw()

    def setup_ui(self):
        main_frame = tk.Frame(self.root, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True)

        control_frame = tk.Frame(main_frame, bg=PANEL_BG, width=430)
        control_frame.pack(side=tk.RIGHT, fill=tk.Y)
        control_frame.pack_propagate(False) 

        # Título Dinâmico
        title_text = I18N[self.lang].get("settings", "⚙️ CONFIGURAÇÕES")
        self.lbl_settings_title = tk.Label(control_frame, text=title_text, bg=PANEL_BG, fg="white", font=("Consolas", 12, "bold"))
        self.lbl_settings_title.pack(pady=(20, 10))

        # Seleção de Idioma
        lang_frame = tk.Frame(control_frame, bg=PANEL_BG)
        lang_frame.pack(fill=tk.X, padx=15, pady=5)
        
        tk.Radiobutton(lang_frame, text="🇺🇸 English", variable=self.lang_var, value="EN", 
                       command=self.on_lang_change, bg=PANEL_BG, fg="white", selectcolor=BTN_BG, 
                       font=("Consolas", 10)).pack(side=tk.LEFT, expand=True)
        
        tk.Radiobutton(lang_frame, text="🇧🇷 Português", variable=self.lang_var, value="PT", 
                       command=self.on_lang_change, bg=PANEL_BG, fg="white", selectcolor=BTN_BG, 
                       font=("Consolas", 10)).pack(side=tk.LEFT, expand=True)

        self.rb_auto = tk.Radiobutton(control_frame, text=I18N[self.lang]["modo_auto"], variable=self.mode, value="AUTO", 
                       command=self.on_mode_change, bg=PANEL_BG, fg="white", selectcolor=BTN_BG, 
                       font=("Consolas", 10))
        self.rb_auto.pack(anchor="w", padx=15, pady=5)
                       
        self.rb_manual = tk.Radiobutton(control_frame, text=I18N[self.lang]["modo_manual"], variable=self.mode, value="MANUAL", 
                       command=self.on_mode_change, bg=PANEL_BG, fg="white", selectcolor=BTN_BG, 
                       font=("Consolas", 10))
        self.rb_manual.pack(anchor="w", padx=15, pady=5)

        self.btn_change_folder = tk.Button(control_frame, text=I18N[self.lang]["btn_change_folder"], command=self.change_directory, 
                                  bg="#555555", fg="white", font=("Consolas", 9), relief=tk.FLAT)
        self.btn_change_folder.pack(pady=(15, 5), padx=15, fill=tk.X)

        # ==========================================
        # INTEGRAÇÃO EXCEL (CONDICIONAL)
        # ==========================================
        # Dois botões independentes lado a lado, cada um com uma única
        # responsabilidade (nunca um botão fazendo duas coisas):
        # - btn_sync_talentos: só sincroniza Talento/Level na planilha principal.
        # - btn_sync_comparador: só roda a aba "Comparações_Talento".
        # Cada um abre/usa sua própria conexão com o Excel, então clicar num
        # não espera nem atrapalha o outro.
        sync_buttons_frame = tk.Frame(control_frame, bg=PANEL_BG)

        self.btn_sync_talentos = tk.Button(
            sync_buttons_frame, text=I18N[self.lang]["btn_sync_talentos"], command=self.on_sync_talentos_click,
            bg="#1F6FEB", fg="white", font=("Consolas", 9, "bold"), relief=tk.FLAT
        )
        self.btn_sync_comparador = tk.Button(
            sync_buttons_frame, text=I18N[self.lang]["btn_sync_comparador"], command=self.on_sync_comparador_click,
            bg="#8A2BE2", fg="white", font=("Consolas", 9, "bold"), relief=tk.FLAT
        )
        auto_sync_frame = tk.Frame(control_frame, bg=PANEL_BG)
        self.rb_auto_sync_on = tk.Radiobutton(
            auto_sync_frame, text=I18N[self.lang]["auto_sync_on"], variable=self.auto_sync_talentos_var,
            value="ON", command=self.on_auto_sync_talentos_change, bg=PANEL_BG, fg="white",
            selectcolor=BTN_BG, activebackground=PANEL_BG, activeforeground="white",
            font=("Consolas", 9), cursor="hand2"
        )
        self.rb_auto_sync_off = tk.Radiobutton(
            auto_sync_frame, text=I18N[self.lang]["auto_sync_off"], variable=self.auto_sync_talentos_var,
            value="OFF", command=self.on_auto_sync_talentos_change, bg=PANEL_BG, fg="white",
            selectcolor=BTN_BG, activebackground=PANEL_BG, activeforeground="white",
            font=("Consolas", 9), cursor="hand2"
        )

        self.lbl_sync_status = tk.Label(control_frame, text="", bg=PANEL_BG, fg=FG_COLOR,
                                         font=("Consolas", 9, "bold"))
        self._sync_status_after_id = None

        # Os botões e o status só ganham layout (pack) se a flag estiver True
        if EXCEL_SYNC_ENABLED:
            sync_buttons_frame.pack(pady=(0, 2), padx=15, fill=tk.X)
            self.btn_sync_talentos.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 3))
            if COMPARADOR_SYNC:
                self.btn_sync_comparador.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(3, 0))
            if AUTO_SYNC_TALENTOS_VISIBLE:
                auto_sync_frame.pack(pady=(0, 2), padx=15, fill=tk.X)
                self.rb_auto_sync_on.pack(side=tk.LEFT, expand=True)
                self.rb_auto_sync_off.pack(side=tk.LEFT, expand=True)
            self.lbl_sync_status.pack(pady=(0, 5), padx=15, fill=tk.X)
        # ==========================================

        self.lbl_inspect = tk.Label(control_frame, text=I18N[self.lang]["lbl_inspect"], bg=PANEL_BG, fg="white", font=("Consolas", 10))
        self.lbl_inspect.pack(anchor="w", padx=15, pady=(15, 5))
        
        self.save_combo = ttk.Combobox(control_frame, state="readonly", width=20)
        self.save_combo.pack(padx=15, pady=5)
        self.save_combo.bind("<<ComboboxSelected>>", self.on_combo_select)
        self.save_combo.bind("<Button-1>", self.update_combo_list)

        self.pause_frame = tk.Frame(control_frame, bg=PANEL_BG)
        self.pause_frame.pack(pady=15, fill=tk.X, padx=15)
        
        self.lbl_paused = tk.Label(self.pause_frame, text=I18N[self.lang]["lbl_paused"], bg=PANEL_BG, fg="yellow", font=("Consolas", 12, "bold"))
        self.btn_resume = tk.Button(self.pause_frame, text=I18N[self.lang]["btn_resume"], command=self.resume_tracking, 
                                    bg="#8B0000", fg="white", font=("Consolas", 10, "bold"), relief=tk.FLAT)

        # ==========================================
        # PAINEL DA WISHLIST (NA BARRA LATERAL)
        # ==========================================
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=(15, 8), padx=15)
        
        self.lbl_wishlist_title = tk.Label(control_frame, text=I18N[self.lang]["wishlist_panel_title"], bg=PANEL_BG, fg="white", font=("Consolas", 10, "bold"))
        self.lbl_wishlist_title.pack(pady=(0, 6))

        # Aviso contextual: só aparece quando pausado/inspecionando um save específico,
        # avisando que dá pra adicionar Digimons desse save na Wishlist aqui embaixo.
        self.lbl_wishlist_context = tk.Label(control_frame, text="", bg=PANEL_BG, fg="#87CEFA",
                                              font=("Consolas", 8, "italic"), wraplength=380, justify="left")
        # Só é "packado" (mostrado) quando há algo pra dizer — ver update_wishlist_context_label()

        # 1. Busca por ID ou Nome (label em cima, campo embaixo, botão de largura total)
        self.lbl_wishlist_search = tk.Label(control_frame, text=I18N[self.lang]["wishlist_search_label"], bg=PANEL_BG, fg="white", font=("Consolas", 9))
        self.lbl_wishlist_search.pack(anchor="w", padx=15, pady=(0, 2))

        self.entry_wish_id = tk.Entry(control_frame, bg="#333333", fg="white", font=("Consolas", 9), relief=tk.FLAT, insertbackground="white")
        self.entry_wish_id.pack(fill='x', padx=15, pady=(0, 4), ipady=2)

        # Autocomplete flutuante: dropdown próprio que aparece embaixo do campo enquanto
        # digita, sem nunca tirar o foco do campo de texto (ver os métodos on_wishlist_query_*).
        self._wish_ac_popup = None
        self._wish_ac_listbox = None
        self._wish_ac_index = -1
        self._wish_ac_results = []
        self._wish_ac_options = []

        self.entry_wish_id.bind("<KeyRelease>", self.on_wishlist_query_keyrelease)
        self.entry_wish_id.bind("<Down>", self.on_wishlist_query_down)
        self.entry_wish_id.bind("<Up>", self.on_wishlist_query_up)
        self.entry_wish_id.bind("<Return>", self.on_wishlist_query_return)
        self.entry_wish_id.bind("<Escape>", self.on_wishlist_query_escape)
        self.entry_wish_id.bind("<Tab>", self.on_wishlist_tab_to_combo)
        # Clicar fora do campo fecha o dropdown (pequeno atraso pra dar tempo do clique
        # num item da lista ser processado antes de fecharmos ela).
        self.entry_wish_id.bind("<FocusOut>", lambda event: self.root.after(150, self.close_wish_autocomplete))

        self.btn_wish_search = tk.Button(control_frame, text=I18N[self.lang]["wishlist_search_btn"], command=self.search_wishlist_digimon, bg="#555555", fg="white", font=("Consolas", 9, "bold"), relief=tk.FLAT, cursor="hand2")
        self.btn_wish_search.pack(fill='y', padx=15, pady=(5, 6), ipady=2)

        choose_header = tk.Frame(control_frame, bg=PANEL_BG)
        choose_header.pack(fill='x', padx=15, pady=(30, 2))

        self.lbl_wishlist_choose = tk.Label(choose_header, text=I18N[self.lang]["wishlist_choose_label"], bg=PANEL_BG, fg="white", font=("Consolas", 9))
        self.lbl_wishlist_choose.pack(side=tk.LEFT)

        self.lbl_wishlist_result_count = tk.Label(choose_header, text="", bg=PANEL_BG, fg="#BBBBBB", font=("Consolas", 9))
        self.lbl_wishlist_result_count.pack(side=tk.RIGHT)
        
        # 2. Combobox de Resultados Encontrados
        self.combo_wish_results = ttk.Combobox(control_frame, state="readonly", font=("Consolas", 9), width=58)
        self.combo_wish_results.pack(fill='x', padx=15, pady=(0, 6))

        # 3. Target Level + Botão Adicionar (lado a lado, já que o campo é só um número de 2 dígitos)
        target_frame = tk.Frame(control_frame, bg=PANEL_BG)
        target_frame.pack(fill='x', padx=15, pady=(30, 10))

        self.lbl_wishlist_target = tk.Label(target_frame, text=I18N[self.lang]["wishlist_target_label"], bg=PANEL_BG, fg="white", font=("Consolas", 9))
        self.lbl_wishlist_target.pack(side=tk.LEFT)

        # Validação: só aceita dígitos e no máximo 2 caracteres (nível vai de 1 a 99)
        vcmd_target = (self.root.register(self._validate_target_lvl_input), '%P')
        self.entry_wish_target = tk.Entry(target_frame, bg="#333333", fg="white", font=("Consolas", 9), relief=tk.FLAT,
                                           insertbackground="white", width=4,
                                           validate="key", validatecommand=vcmd_target)
        self.entry_wish_target.pack(side=tk.LEFT, padx=(5, 5), ipady=2)

        # ATALHO: Aperta Enter no Target Lv e já adiciona a meta
        self.entry_wish_target.bind("<Return>", lambda event: self.add_wishlist_target())

        self.btn_wish_add = tk.Button(target_frame, text=I18N[self.lang]["wishlist_add_btn"], command=self.add_wishlist_target, bg="#005f87", fg="white", font=("Consolas", 9, "bold"), relief=tk.FLAT, cursor="hand2")
        self.btn_wish_add.pack(side=tk.LEFT, fill='x', expand=True, ipady=2)
        # ==========================================

        # ==========================================
        # PAINEL DE RESUMO (NA BARRA LATERAL)
        # ==========================================
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=(15, 8), padx=15)

        self.lbl_summary_title = tk.Label(control_frame, text=I18N[self.lang]["lbl_summary"], bg=PANEL_BG, fg="white", font=("Consolas", 10, "bold"))
        self.lbl_summary_title.pack(pady=(0, 6))

        self.summary_text = tk.Text(control_frame, bg=BG_COLOR, font=("Consolas", 9),
                                     wrap=tk.WORD, state=tk.DISABLED, bd=0, height=6,
                                     padx=8, pady=6)
        self.summary_text.pack(fill='x', padx=15, pady=(0, 10))

        self.summary_text.tag_config("alert", foreground=FG_ALERT, font=("Consolas", 9, "bold"))
        self.summary_text.tag_config("almost", foreground=FG_ALMOST)
        self.summary_text.tag_config("status", foreground="white")
        self.summary_text.tag_config("loc_fazenda", foreground="#32CD32")

        # ------------------------------------------
        # Painel de ocupação (Box / Fazenda), separado, com fonte maior
        # ------------------------------------------
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=(0, 8), padx=15)

        occupancy_frame = tk.Frame(control_frame, bg=PANEL_BG)
        occupancy_frame.pack(fill='x', padx=15, pady=(0, 10))

        self.lbl_party_count = tk.Label(occupancy_frame, text="PARTY: -/6", bg=PANEL_BG, fg="#FFA500", font=("Consolas", 13, "bold"))
        self.lbl_party_count.pack(side=tk.LEFT, padx=(0, 12))

        self.lbl_box_count = tk.Label(occupancy_frame, text="BOX: -/999", bg=PANEL_BG, fg="#1E90FF", font=("Consolas", 13, "bold"))
        self.lbl_box_count.pack(side=tk.LEFT, padx=(0, 12))

        self.lbl_farm_count = tk.Label(occupancy_frame, text="FAZENDA: -/30", bg=PANEL_BG, fg="#32CD32", font=("Consolas", 13, "bold"))
        self.lbl_farm_count.pack(side=tk.LEFT)

        self.lbl_total_count = tk.Label(control_frame, text="TOTAL: -", bg=PANEL_BG, fg="white", font=("Consolas", 12, "bold"))
        self.lbl_total_count.pack(anchor="w", padx=15, pady=(8, 0))
        # ==========================================

        log_frame = tk.Frame(main_frame, bg=BG_COLOR)
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(log_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_area = tk.Text(log_frame, bg=BG_COLOR, fg=FG_COLOR, font=("Consolas", 11), 
                                 wrap=tk.WORD, state=tk.DISABLED, bd=0, padx=20, pady=20,
                                 yscrollcommand=self.scrollbar.set, spacing3=3)
        self.text_area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar.config(command=self.text_area.yview)
        
        self.text_area.tag_config("alert", foreground=FG_ALERT, font=("Consolas", 11, "bold"))
        self.text_area.tag_config("almost", foreground=FG_ALMOST)
        self.text_area.tag_config("header", foreground="#00BFFF", font=("Consolas", 12, "bold"))
        self.text_area.tag_config("header_red", foreground="#FF4D4D", font=("Consolas", 12, "bold"))
        self.text_area.tag_config("header_orange", foreground="#FFA500", font=("Consolas", 12, "bold"))
        self.text_area.tag_config("header_yellow", foreground="#FFD700", font=("Consolas", 12, "bold"))
        self.text_area.tag_config("header_green", foreground="#32CD32", font=("Consolas", 12, "bold"))
        self.text_area.tag_config("status", foreground="white")
        self.text_area.tag_config("loc_party", foreground="#FFA500") 
        self.text_area.tag_config("loc_box", foreground="#1E90FF")   
        self.text_area.tag_config("loc_fazenda", foreground="#32CD32") 

        self.quase_filter_loc_var = tk.StringVar(value=self.quase_filter_location)
        self.quase_filter_protected_var = tk.BooleanVar(value=self.quase_filter_protected)
        self.quase_filter_wishlist_var = tk.BooleanVar(value=self.quase_filter_wishlist)
        self.quase_limit_entry_var = tk.StringVar()
        self.quase_search_var = tk.StringVar()
        self.quase_search_text = ""  # busca ao vivo da Quase Lá — só de sessão, não persiste no config.json
        self._quase_search_after_id = None
        self._quase_wishlist_hint_text = I18N[self.lang]["checkbox_wishlist_help"]

        self.update_wishlist_context_label()

    def _validate_target_lvl_input(self, proposed_value):
        """Validação de campo (usada pelo Entry do Target Level): só aceita vazio ou até 2 dígitos numéricos (nível vai de 1 a 99)."""
        if proposed_value == "":
            return True
        return proposed_value.isdigit() and len(proposed_value) <= 2

    def read_digimon_uid(self, data, name_offset):
        """UID do individuo no mesmo formato visto no HxD, relativo ao inicio do nome."""
        uid_start = name_offset - 0x08
        uid_end = name_offset - 0x04
        if uid_start < 0 or uid_end > len(data):
            return ""
        return data[uid_start:uid_end].hex().upper()

    def read_digimon_protected(self, data, name_offset):
        """Lê o status de proteção (cadeado) do Digimon. Offset achado via Cheat Engine:
        Protection = Talent (name_offset + 0x100) + 0x1C = name_offset + 0x11C.
        Utiliza máscara de bits (& 1) para capturar o cadeado mesmo se o Digimon
        tiver outros status (como 256) somados no mesmo byte."""
        offset = name_offset + 0x11C
        if offset + 4 > len(data):
            return False
        
        # Lê o valor bruto salvo na memória (ex: 0, 1, 256, 257)
        status_value = struct.unpack_from("<I", data, offset)[0]
        
        # Retorna True apenas se o bit correspondente ao valor 1 estiver ativo
        return (status_value & 1) != 0
    
    def _connect_to_excel_app_and_book(self):
        """
        Acha (ou abre) o app/book do Excel referente a EXCEL_FILE_PATH.
        Compartilhado pelos dois botões de sync (Talentos e Comparador),
        já que ambos precisam da mesma lógica de conexão. Levanta exceção
        se não achar/abrir o arquivo - quem chama deve envolver em try/except.
        """
        app = None
        book = None
        target_name = os.path.basename(EXCEL_FILE_PATH).lower()
        for running_app in xw.apps:
            for bk in running_app.books:
                if bk.name.lower() == target_name:
                    app = running_app
                    book = bk
                    break
            if book:
                break

        if book is None:
            app = xw.apps.active if xw.apps else xw.App(visible=True)
            book = app.books.open(EXCEL_FILE_PATH)

        return app, book

    def sync_protected_talents_to_excel(self, active_digimons):
        """
        Sincroniza o ascendant_talent_raw e o level dos Digimons protegidos com a
        planilha timestranger.xlsm, mantendo o Excel aberto (via xlwings).

        Só mexe na aba principal (EXCEL_SHEET_NAME) - a aba "Comparações_Talento"
        tem seu próprio botão/método (sync_comparador_to_excel), pra um botão
        nunca fazer duas coisas ao mesmo tempo e pra cada sync ser o mais leve
        e rápido possível no que faz.

        IMPORTANTE: este método é chamado sob demanda (botão "🔄 Sync Main Sheet"),
        rodando numa thread separada da UI - por isso ele NÃO toca em nenhum
        widget do tkinter diretamente (nem self.log). Em vez disso, devolve
        um dict de resultado; quem atualiza a tela é o handler do botão
        (on_sync_talentos_click), via self.root.after(...).

        Performático porque:
        - Só roda quando o usuário pede (não mais a cada leitura de save,
          que era o gargalo antes).
        - Lê a Coluna F inteira de uma vez (1 chamada COM) e monta um dict
          nome -> linha em memória, em vez de buscar célula por célula.
        - Escreve em lote (um único range 2D), em vez de 1 write por Digimon.
        - Desliga screen_updating/cálculo/eventos automáticos durante a
          escrita e restaura no final, mesmo se der erro.
        - Mira sempre a aba EXCEL_SHEET_NAME (nome fixo), não a aba que
          estiver visível/ativa no Excel no momento do clique.

        Retorna: {'status': 'disabled'|'no_lib'|'no_protected'|'no_match'|'ok'|'error',
                  'count': int, 'error': str|None}
        """
        if not EXCEL_SYNC_ENABLED:
            return {'status': 'disabled', 'count': 0, 'error': None}
        if not XLWINGS_AVAILABLE:
            return {'status': 'no_lib', 'count': 0, 'error': None}

        protegidos = [d for d in active_digimons.values() if d.get('protected')]
        if not protegidos:
            return {'status': 'no_protected', 'count': 0, 'error': None}

        app = None
        prev_screen_updating = None
        prev_calculation = None
        prev_enable_events = None

        try:
            app, book = self._connect_to_excel_app_and_book()
            sheet = book.sheets[EXCEL_SHEET_NAME] if EXCEL_SHEET_NAME else book.sheets.active

            prev_screen_updating = app.screen_updating
            prev_calculation = app.calculation
            prev_enable_events = app.api.EnableEvents
            app.screen_updating = False
            app.calculation = 'manual'
            # Desliga eventos do Excel (ex.: macros de Worksheet_Change) durante
            # a escrita. Num .xlsm com macros, cada célula escrita pode disparar
            # uma macro e isso é uma causa clássica de sync lento.
            app.api.EnableEvents = False

            used_range = sheet.used_range
            last_row = used_range.last_cell.row

            main_status = 'ok'
            main_count = 0

            if last_row < EXCEL_HEADER_ROW:
                main_status = 'no_match'
            else:
                col_values = sheet.range(
                    (EXCEL_HEADER_ROW, EXCEL_COL_NAME),
                    (last_row, EXCEL_COL_NAME)
                ).value
                if not isinstance(col_values, list):
                    col_values = [col_values]

                name_to_row = {}
                for idx, val in enumerate(col_values):
                    if val:
                        name_to_row[str(val).strip().upper()] = EXCEL_HEADER_ROW + idx

                # Armazena tanto o talento quanto o level para cada linha correspondente
                row_updates = {}
                for dig in protegidos:
                    name_key = str(dig.get('name', '')).strip().upper()
                    row = name_to_row.get(name_key)
                    if row:
                        row_updates[row] = [
                            dig.get('ascendant_talent'),
                            dig.get('level'),
                            dig.get('elo')
                        ]

                if not row_updates:
                    main_status = 'no_match'
                else:
                    # OTIMIZAÇÃO ANTERIOR (removida): lia o bloco inteiro
                    # (1ª à última linha protegida), alterava em memória só
                    # as linhas dos protegidos e escrevia o bloco inteiro de
                    # volta. Rápido, mas ARRISCADO com filtro ativo na
                    # planilha: esse bloco cobre também linhas de Digimons
                    # NÃO protegidos entre um protegido e outro, e ranges do
                    # Excel que cruzam linhas ocultas por filtro podem se
                    # comportar de forma inconsistente ao ler/escrever esse
                    # tipo de bloco "misto" - o que podia acabar gravando
                    # valor errado em linha errada bem no meio do bloco.
                    #
                    # CORREÇÃO: nunca lemos valor nenhum de volta. Só
                    # escrevemos, linha por linha protegida, os valores que
                    # JÁ sabemos que são os certos (vieram do save, não da
                    # planilha). Isso é 100% imune a filtro, porque cada
                    # escrita mira exatamente a linha certa (endereço
                    # absoluto, vindo do name_to_row lido agora mesmo) e
                    # nunca depende do que está nas linhas vizinhas.
                    #
                    # Ainda assim continua rápido: agrupamos linhas
                    # protegidas CONSECUTIVAS num único range de escrita
                    # (Talento + Level + Elo juntos), em vez de 1 chamada
                    # COM por Digimon.
                    rows_sorted = sorted(row_updates.keys())

                    block_start = rows_sorted[0]
                    block_values = [row_updates[rows_sorted[0]]]
                    prev_row = rows_sorted[0]

                    def _flush_block(start_row, values_list):
                        end_row = start_row + len(values_list) - 1
                        sheet.range(
                            (start_row, EXCEL_COL_ASCENDANT),
                            (end_row, EXCEL_COL_ELO)
                        ).value = values_list

                    for row in rows_sorted[1:]:
                        if row == prev_row + 1:
                            block_values.append(row_updates[row])
                        else:
                            _flush_block(block_start, block_values)
                            block_start = row
                            block_values = [row_updates[row]]
                        prev_row = row
                    _flush_block(block_start, block_values)

                    main_count = len(row_updates)

            return {'status': main_status, 'count': main_count, 'error': None}

        except Exception as e:
            return {'status': 'error', 'count': 0, 'error': str(e)}
        finally:
            if app is not None:
                try:
                    if prev_calculation is not None:
                        app.calculation = prev_calculation
                    if prev_screen_updating is not None:
                        app.screen_updating = prev_screen_updating
                    if prev_enable_events is not None:
                        app.api.EnableEvents = prev_enable_events
                except Exception:
                    pass

    def sync_comparador_to_excel(self, active_digimons):
        """
        Roda SÓ a aba "Comparações_Talento" (histórico/comparação de Talento
        Ascendente) - não toca na planilha principal. Botão separado
        (btn_sync_comparador), conexão própria com o Excel, independente do
        sync de Talentos/Level.

        Chamado numa thread separada; não toca em widgets do tkinter
        diretamente (mesma lógica de segurança de thread do outro sync).

        Retorna o dict de status de _sync_comparison_sheet diretamente:
        {'status': 'created'|'updated'|'reset'|'disabled'|'no_lib'|'no_protected'|'error',
         'count': int, 'new': int, 'error': str|None}
        """
        if not EXCEL_SYNC_ENABLED:
            return {'status': 'disabled', 'count': 0, 'new': 0, 'error': None}
        if not COMPARADOR_SYNC:
            return {'status': 'disabled', 'count': 0, 'new': 0, 'error': None}
        if not XLWINGS_AVAILABLE:
            return {'status': 'no_lib', 'count': 0, 'new': 0, 'error': None}

        protegidos = [d for d in active_digimons.values() if d.get('protected')]
        if not protegidos:
            return {'status': 'no_protected', 'count': 0, 'new': 0, 'error': None}

        app = None
        prev_screen_updating = None
        prev_calculation = None
        prev_enable_events = None

        try:
            app, book = self._connect_to_excel_app_and_book()

            prev_screen_updating = app.screen_updating
            prev_calculation = app.calculation
            prev_enable_events = app.api.EnableEvents
            app.screen_updating = False
            app.calculation = 'manual'
            app.api.EnableEvents = False

            return self._sync_comparison_sheet(book, protegidos)

        except Exception as e:
            return {'status': 'error', 'count': 0, 'new': 0, 'error': str(e)}
        finally:
            if app is not None:
                try:
                    if prev_calculation is not None:
                        app.calculation = prev_calculation
                    if prev_screen_updating is not None:
                        app.screen_updating = prev_screen_updating
                    if prev_enable_events is not None:
                        app.api.EnableEvents = prev_enable_events
                except Exception:
                    pass

    def _sync_comparison_sheet(self, book, protegidos):
        """
        Popula/atualiza a aba "Comparações_Talento" com o histórico.
        Trabalha em Ciclos de 5 rodadas: ao chegar na 5ª, ele reinicia
        a tabela transformando o resultado na nova Coluna A e B.
        Ordena automaticamente do maior ganho para o menor.
        
        Popula/atualiza a aba "Comparações_Talento" mantendo a estrutura limpa.
        - Preserva as cores (azul/vermelho) de TODAS as colunas de comparação.
        - Ordena por:
          1ª Prioridade: Aumento de talento (maior aumento primeiro)
          2ª Prioridade: Menor talento atual após a comparação
          3ª Prioridade: Ordem alfabética
        

        Chamada de dentro de sync_protected_talents_to_excel, reaproveitando
        o MESMO `book` já aberto (não abre uma conexão nova com o Excel).

        Layout da aba (colunas crescem pra direita a cada rodada de sync):
            A: "[ Local ] Nome do Digimon"   <- chave da linha (nome, sem colchetes)
            B: Talento Inicial               <- 1ª extração (baseline, sem comparação)
            C: Comparação 2   D: Talento 2   <- a partir da 2ª extração: par
            E: Comparação 3   F: Talento 3      (Comparação, Talento) por rodada
            ...

        Comparação (colunas C, E, G, ...):
            "[ Local Atual ] Nome Aumentou +X!"   -> fonte azul
            "[ Local Atual ] Nome Diminuiu -X!"   -> fonte vermelha
            "[ Local Atual ] Nome Mesmo talento!" -> sem cor

        Retorna: {'status': 'created'|'updated'|'error', 'count': int, 'new': int, 'error': str|None}
        """
        try:
            sheet_names = [s.name for s in book.sheets]

            # -------------------------------------------------------------
            # FUNÇÃO AUXILIAR PARA CRIAR O BASELINE (Aba inexistente ou vazia)
            # -------------------------------------------------------------
            def _create_baseline(sheet_obj):
                sheet_obj.clear()
                rows = []
                for dig in protegidos:
                    local = self._loc_display(dig.get('loc', '?'), dig.get('slot'), translate=False)
                    name = str(dig.get('name', '')).strip()
                    talent = dig.get('ascendant_talent')
                    rows.append({
                        'label': f"[ {local} ] {name}",
                        'name': name,
                        'talent': talent if isinstance(talent, (int, float)) else 999999999
                    })
                
                # Ordena baseline por menor talento, depois alfabético
                rows.sort(key=lambda x: (x['talent'], x['name'].upper()))

                grid = [["Digimon", "Talento Inicial"]]
                for r in rows:
                    t_val = r['talent'] if r['talent'] != 999999999 else ""
                    grid.append([r['label'], t_val])

                sheet_obj.range((1, 1)).value = grid
                sheet_obj.autofit()
                try:
                    sheet_obj.api.AutoFilterMode = False
                    sheet_obj.range("A1:B1").api.AutoFilter(1)
                except Exception: pass
                return {'status': 'created', 'count': len(rows), 'new': 0, 'error': None}

            # CASO 1: Aba não existe
            if EXCEL_COMPARISON_SHEET_NAME not in sheet_names:
                sheet = book.sheets.add(EXCEL_COMPARISON_SHEET_NAME, after=book.sheets[-1])
                return _create_baseline(sheet)

            sheet = book.sheets[EXCEL_COMPARISON_SHEET_NAME]
            used_range = sheet.used_range
            last_row = used_range.last_cell.row
            last_col = used_range.last_cell.column

            if last_row < 2:
                return _create_baseline(sheet)

            # Lê a matriz completa de dados atuais da planilha
            full_data = sheet.range((1, 1), (last_row, last_col)).value
            if not isinstance(full_data, list):
                full_data = [[full_data]]
            elif full_data and not isinstance(full_data[0], list):
                full_data = [full_data]

            headers = full_data[0]
            
            # Identifica os índices das colunas de Talento (0-based)
            talent_cols_indices = [idx for idx, h in enumerate(headers) if h and str(h).strip().lower().startswith("talento")]

            rounds_so_far = len(talent_cols_indices)
            last_talent_idx = talent_cols_indices[-1] if talent_cols_indices else 1
            
            name_pattern = re.compile(r'^\[\s*[^\]]*\]\s*(.*)$')

            # -------------------------------------------------------------
            # REINÍCIO DE CICLO (RODADA MÁXIMA -> RESET VISUAL PARA COLUNAS A E B)
            # -------------------------------------------------------------
            if rounds_so_far >= MAX_COMPARACOES:
                # Transfere APENAS os dados da última rodada registrada na planilha
                # para a Coluna A e B. Ignora a extração atual (protegidos) para não 
                # perder a visualização de comparação. O usuário clica de novo após o reset.
                sheet.clear()

                # Índice nome_puro -> Digimon protegido NESTA extração, pra
                # conseguir atualizar o [ Local ] de cada linha pro local
                # ATUAL (o antigo "[ Local ] Nome" da Coluna A guarda o
                # local de quando aquela linha foi escrita, que pode já
                # estar desatualizado - ex.: Digimon saiu da PARTY e foi
                # pra BOX entre uma sincronização e outra).
                protegido_by_name = {
                    str(d.get('name', '')).strip().upper(): d for d in protegidos
                }

                rows_reset = []
                for row in full_data[1:]:
                    row_cells = list(row)
                    raw_label = row_cells[0] if len(row_cells) > 0 else ""
                    old_talent_val = row_cells[last_talent_idx] if len(row_cells) > last_talent_idx else None
                    
                    if raw_label:
                        m = name_pattern.match(str(raw_label))
                        bare_name = m.group(1).strip() if m else str(raw_label).strip()

                        # Atualiza o [ Local ] pro local ATUAL do Digimon, se
                        # ele ainda estiver protegido nesta extração. Se não
                        # estiver mais protegido/não foi encontrado, mantém
                        # o label antigo como estava (não temos como saber
                        # o local atual dele).
                        dig_atual = protegido_by_name.get(bare_name.upper())
                        if dig_atual is not None:
                            novo_local = self._loc_display(dig_atual.get('loc', '?'), dig_atual.get('slot'), translate=False)
                            label_atualizado = f"[ {novo_local} ] {bare_name}"
                        else:
                            label_atualizado = str(raw_label).strip()

                        rows_reset.append({
                            'label': label_atualizado,
                            'name': bare_name,
                            'talent': old_talent_val if isinstance(old_talent_val, (int, float)) else 999999999
                        })
                
                # Mantém a ordenação: 1º menor talento, 2º ordem alfabética
                rows_reset.sort(key=lambda x: (x['talent'], x['name'].upper()))

                grid_reset = [["Digimon", "Talento Inicial"]]
                for r in rows_reset:
                    t_val = r['talent'] if r['talent'] != 999999999 else ""
                    grid_reset.append([r['label'], t_val])

                sheet.range((1, 1)).value = grid_reset
                sheet.autofit()
                
                try:
                    sheet.api.AutoFilterMode = False
                    sheet.range("A1:B1").api.AutoFilter(1)
                except Exception: pass
                
                # Retorna um status de reset para sinalizar que reorganizou
                return {'status': 'reset', 'count': len(rows_reset), 'new': 0, 'error': None}

            # -------------------------------------------------------------
            # RODADA NORMAL (Adiciona Comparação X e Talento X)
            # -------------------------------------------------------------
            new_round_num = rounds_so_far + 1
            new_comp_header = f"Comparação {new_round_num}"
            new_value_header = f"Talento {new_round_num}"

            new_headers = list(headers) + [new_comp_header, new_value_header]
            
            protegido_by_name = {str(d.get('name', '')).strip().upper(): d for d in protegidos}

            processed_rows = []
            matched_names = set()

            # Processa as linhas existentes na planilha
            for row in full_data[1:]:
                row_cells = list(row)
                raw_name_cell = row_cells[0] if len(row_cells) > 0 else ""
                old_talent_val = row_cells[last_talent_idx] if len(row_cells) > last_talent_idx else None

                bare_name = ""
                if raw_name_cell:
                    m = name_pattern.match(str(raw_name_cell))
                    bare_name = m.group(1).strip() if m else str(raw_name_cell).strip()

                key = bare_name.upper()
                dig = protegido_by_name.get(key)

                if dig is None:
                    comp_text = f"[ ? ] {bare_name} - Digimon não encontrado nesta rodada"
                    new_talent = ""
                    diff = -999999
                else:
                    matched_names.add(key)
                    local = self._loc_display(dig.get('loc', '?'), dig.get('slot'), translate=False)
                    new_talent = dig.get('ascendant_talent')
                    
                    diff = 0
                    if isinstance(old_talent_val, (int, float)) and isinstance(new_talent, (int, float)):
                        diff = new_talent - old_talent_val
                        if diff > 0:
                            status_str = f"Aumentou +{int(diff)}!"
                        elif diff < 0:
                            status_str = f"Diminuiu -{abs(int(diff))}!"
                        else:
                            status_str = "Mesmo talento!"
                    else:
                        status_str = "Sem valor anterior para comparar."

                    comp_text = f"[ {local} ] {bare_name} - {status_str}"

                # Anexa as duas novas células ao final da linha existente
                row_cells.append(comp_text)
                row_cells.append(new_talent)

                processed_rows.append({
                    'row_cells': row_cells,
                    'bare_name': bare_name,
                    'diff': diff,
                    'new_talent': new_talent
                })

            # Adiciona Digimons novos que surgiram nesta rodada
            novos = [d for d in protegidos if str(d.get('name', '')).strip().upper() not in matched_names]
            for dig in novos:
                local = self._loc_display(dig.get('loc', '?'), dig.get('slot'), translate=False)
                bare_name = str(dig.get('name', '')).strip()
                new_talent = dig.get('ascendant_talent')

                new_row_cells = [f"[ {local} ] {bare_name}"]
                # Preenche com colunas vazias até atingir o número de colunas antigas
                while len(new_row_cells) < len(headers):
                    new_row_cells.append("")
                
                comp_text = f"[ {local} ] {bare_name} - Novo Digimon protegido - sem comparação anterior."
                new_row_cells.append(comp_text)
                new_row_cells.append(new_talent)

                processed_rows.append({
                    'row_cells': new_row_cells,
                    'bare_name': bare_name,
                    'diff': 0,
                    'new_talent': new_talent
                })

            # -------------------------------------------------------------
            # ORDENAÇÃO SEGUNDO AS TRÊS PRIORIDADES:
            # 1ª Prioridade: Digimons que aumentaram talento (maior ganho no topo)
            # 2ª Prioridade: Menor talento após a comparação
            # 3ª Prioridade: Ordem alfabética
            # -------------------------------------------------------------
            def sort_key(item):
                d = item['diff']
                t_val = item['new_talent'] if isinstance(item['new_talent'], (int, float)) else 999999999
                name = item['bare_name'].upper()

                if d > 0:
                    return (0, -d, t_val, name)
                else:
                    return (1, 0, t_val, name)

            processed_rows.sort(key=sort_key)

            # Monta a matriz final
            final_grid = [new_headers]
            for r in processed_rows:
                final_grid.append(r['row_cells'])

            # Reescreve a planilha
            sheet.clear()
            sheet.range((1, 1)).value = final_grid

            # -------------------------------------------------------------
            # PRESERVAÇÃO DE CORES EM TODAS AS COLUNAS DE COMPARAÇÃO
            #
            # OTIMIZAÇÃO DE PERFORMANCE: sheet.range(...).font.color = ... é
            # 1 chamada COM (ida-e-volta pro processo do Excel) POR CÉLULA.
            # Com várias colunas "Comparação" x centenas de Digimons, um
            # loop célula-a-célula facilmente vira milhares de chamadas e
            # trava o programa por vários segundos.
            #
            # Em vez disso, agrupamos linhas CONSECUTIVAS que precisam da
            # MESMA cor em blocos, e aplicamos a cor com 1 única chamada
            # por bloco (sheet.range((linha_ini,col),(linha_fim,col)).font.color).
            # Isso reduz de "1 chamada por Digimon" pra "1 chamada por
            # sequência contínua da mesma cor" - normalmente uma fração
            # pequena do total.
            # -------------------------------------------------------------
            COLOR_BLUE = (30, 100, 255)
            COLOR_RED = (220, 30, 30)

            for col_idx, h_name in enumerate(new_headers, start=1):
                if not (h_name and str(h_name).strip().lower().startswith("comparação")):
                    continue

                # Monta a lista de cor por linha (None = sem cor / não mexe)
                row_colors = []
                for row_i in range(2, len(final_grid) + 1):
                    cell_val = str(final_grid[row_i - 1][col_idx - 1] or "").lower()
                    if "aumentou" in cell_val:
                        row_colors.append((row_i, COLOR_BLUE))
                    elif "diminuiu" in cell_val:
                        row_colors.append((row_i, COLOR_RED))
                    # "Mesmo talento!" ou vazio -> sem cor, não precisa de chamada nenhuma

                if not row_colors:
                    continue

                # Agrupa linhas consecutivas com a MESMA cor em blocos
                block_start_row, block_color = row_colors[0]
                prev_row = block_start_row
                for row_i, color in row_colors[1:]:
                    if row_i == prev_row + 1 and color == block_color:
                        prev_row = row_i
                        continue
                    # Fecha o bloco atual (1 chamada COM pro bloco inteiro)
                    sheet.range((block_start_row, col_idx), (prev_row, col_idx)).font.color = block_color
                    block_start_row, block_color = row_i, color
                    prev_row = row_i
                # Fecha o último bloco pendente
                sheet.range((block_start_row, col_idx), (prev_row, col_idx)).font.color = block_color

            # OTIMIZAÇÃO DE PERFORMANCE: sheet.autofit() sem escopo mede o
            # texto de TODAS as colunas da aba (e esse conjunto só cresce a
            # cada rodada). sheet.clear() NÃO reseta a largura de coluna já
            # calculada em rodadas anteriores (largura é propriedade da
            # coluna, não é apagada junto do conteúdo/formatação das
            # células) - então só precisamos ajustar a largura das DUAS
            # colunas novas desta rodada, não da aba inteira.
            new_col_start = len(headers) + 1
            new_col_end = len(new_headers)
            sheet.range((1, new_col_start), (len(final_grid), new_col_end)).autofit()

            # Aplica AutoFiltro na tabela inteira
            try:
                sheet.api.AutoFilterMode = False
                sheet.range((1, 1), (1, len(new_headers))).api.AutoFilter(1)
            except Exception: pass

            return {'status': 'updated', 'count': len(processed_rows), 'new': len(novos), 'error': None}

        except Exception as e:
            return {'status': 'error', 'count': 0, 'new': 0, 'error': str(e)}

    def focus_wishlist_search(self):
        if not hasattr(self, "entry_wish_id"):
            return
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(150, lambda: self.root.attributes("-topmost", False))
            self.root.focus_force()
            self.entry_wish_id.focus_set()
            self.entry_wish_id.icursor(tk.END)
        except Exception:
            pass

    def setup_global_hotkey(self):
        """
        Registra o atalho global (HOTKEY_SYNC_TALENTOS) que dispara o mesmo
        clique do botão azul "🔄 Sync Main Sheet", funcionando mesmo com a
        janela do programa em segundo plano/minimizada.

        Usa a lib "keyboard", que faz um hook de teclado em baixo nível do
        Windows - por isso pega o toque independente de qual janela está
        em foco. Como o callback dela roda numa thread própria (não a
        thread principal do tkinter), a gente não chama on_sync_talentos_click
        direto: agenda com self.root.after(0, ...) pra rodar com segurança
        na thread da UI.
        """
        self._hotkey_registered = False
        self._registered_hotkeys = []

        if not HOTKEY_SYNC_ENABLED:
            return
        if not KEYBOARD_AVAILABLE:
            self.log(f" {I18N[self.lang]['msg_hotkey_no_lib']}", "alert")
            return

        try:
            keyboard.add_hotkey(
                HOTKEY_SYNC_TALENTOS,
                lambda: self.root.after(0, self.on_sync_talentos_click)
            )
            self._registered_hotkeys.append(HOTKEY_SYNC_TALENTOS)
            keyboard.add_hotkey(
                HOTKEY_FOCUS_WISHLIST,
                lambda: self.root.after(0, self.focus_wishlist_search)
            )
            self._registered_hotkeys.append(HOTKEY_FOCUS_WISHLIST)
            self._hotkey_registered = True
            self.log(f" {I18N[self.lang]['msg_hotkey_registered'].format(key=HOTKEY_SYNC_TALENTOS)}", "status")
            self.log(f" {I18N[self.lang]['msg_hotkey_registered'].format(key=HOTKEY_FOCUS_WISHLIST)}", "status")
        except Exception as e:
            for hotkey in getattr(self, "_registered_hotkeys", []):
                try:
                    keyboard.remove_hotkey(hotkey)
                except Exception:
                    pass
            self._registered_hotkeys = []
            self._hotkey_registered = False
            self.log(f" {I18N[self.lang]['msg_hotkey_failed'].format(key=f'{HOTKEY_SYNC_TALENTOS}/{HOTKEY_FOCUS_WISHLIST}')}{e}", "alert")

    def teardown_global_hotkey(self):
        """Desregistra o atalho global ao fechar o programa, evitando o hook
        de teclado ficar "preso" no sistema depois que a janela já fechou."""
        if getattr(self, "_hotkey_registered", False) and KEYBOARD_AVAILABLE:
            for hotkey in getattr(self, "_registered_hotkeys", [HOTKEY_SYNC_TALENTOS]):
                try:
                    keyboard.remove_hotkey(hotkey)
                except Exception:
                    pass

    def _legacy_on_sync_talentos_click(self):
        """
        Handler do botão "🔄 Sync Main Sheet". Só sincroniza Talento/Level na
        planilha principal - não mexe na aba de comparação. Roda numa thread
        separada pra não travar a janela; a UI só é tocada de volta na
        thread principal, via self.root.after(...).
        """
        t = I18N[self.lang]

        if getattr(self, "_sync_talentos_running", False):
            return  # já tem uma sincronização rodando, ignora clique duplicado

        active_digimons = getattr(self, "active_digimons", None)
        if not active_digimons:
            self.log(f" {t['msg_sync_excel_no_data']}", "alert")
            return

        self._sync_talentos_running = True
        self.btn_sync_talentos.config(state=tk.DISABLED, text=t["btn_sync_talentos_running"])
        self._cancel_sync_status_timer()
        self.lbl_sync_status.config(text="")

        def worker():
            result = self.sync_protected_talents_to_excel(active_digimons)
            self.root.after(0, lambda: self._on_sync_talentos_done(result))

        threading.Thread(target=worker, daemon=True).start()

    def on_auto_sync_talentos_change(self):
        self.auto_sync_talentos_enabled = (
            AUTO_SYNC_TALENTOS_VISIBLE and self.auto_sync_talentos_var.get() == "ON"
        )
        self.save_config()
        t = I18N[self.lang]
        msg = t["msg_auto_sync_on"] if self.auto_sync_talentos_enabled else t["msg_auto_sync_off"]
        self._show_sync_status_label(msg, FG_COLOR if self.auto_sync_talentos_enabled else FG_ALMOST)

    def should_auto_sync_talentos(self):
        return (
            EXCEL_SYNC_ENABLED
            and AUTO_SYNC_TALENTOS_VISIBLE
            and getattr(self, "auto_sync_talentos_enabled", False)
        )

    def start_sync_talentos(self, active_digimons=None, silent_no_data=False):
        t = I18N[self.lang]

        if getattr(self, "_sync_talentos_running", False):
            return False

        if active_digimons is None:
            active_digimons = getattr(self, "active_digimons", None)
        if not active_digimons:
            if not silent_no_data:
                self.log(f" {t['msg_sync_excel_no_data']}", "alert")
            return False

        active_snapshot = {key: dict(value) for key, value in active_digimons.items()}
        self._sync_talentos_running = True
        self.btn_sync_talentos.config(state=tk.DISABLED, text=t["btn_sync_talentos_running"])
        self._cancel_sync_status_timer()
        self.lbl_sync_status.config(text="")

        def worker():
            result = self.sync_protected_talents_to_excel(active_snapshot)
            self.root.after(0, lambda: self._on_sync_talentos_done(result))

        threading.Thread(target=worker, daemon=True).start()
        return True

    def maybe_auto_sync_talentos(self, filename, save_mtime):
        if not self.should_auto_sync_talentos():
            return
        signature = (filename, save_mtime)
        if signature == self._last_auto_sync_signature:
            return
        if self.start_sync_talentos(silent_no_data=True):
            self._last_auto_sync_signature = signature

    def on_sync_talentos_click(self):
        self.start_sync_talentos()

    def _on_sync_talentos_done(self, result):
        """Roda na thread principal (chamado via root.after) pra atualizar a UI com segurança."""
        t = I18N[self.lang]
        self._sync_talentos_running = False
        self.btn_sync_talentos.config(state=tk.NORMAL, text=t["btn_sync_talentos"])

        status = result.get('status')
        label_msg, label_color = None, FG_COLOR

        if status == 'ok':
            msg = f"{t['msg_sync_excel_ok']}{result.get('count', 0)}"
            self.log(f" {msg}", "status")
            label_msg, label_color = msg, FG_COLOR
        elif status in ('no_protected', 'no_match'):
            msg = t['msg_sync_excel_none']
            self.log(f" {msg}", "status")
            label_msg, label_color = msg, FG_ALMOST
        elif status == 'disabled':
            msg = t['msg_sync_excel_disabled']
            self.log(f" {msg}", "alert")
            label_msg, label_color = msg, FG_ALERT
        elif status == 'no_lib':
            msg = "⚠️ xlwings não instalado."
            self.log(" ⚠️ xlwings não instalado - sync com Excel desativado.", "alert")
            label_msg, label_color = msg, FG_ALERT
        elif status == 'error':
            self.log(f" ⚠️ [Excel Sync] Falha ao atualizar planilha: {result.get('error')}", "alert")
            label_msg, label_color = "⚠️ Erro ao sincronizar.", FG_ALERT

        if label_msg:
            self._show_sync_status_label(label_msg, label_color)

    def on_sync_comparador_click(self):
        """
        Handler do botão "📊 Sync Comparator". Só roda a aba
        "Comparações_Talento" - não mexe na planilha principal. Roda numa
        thread separada e independente do sync de Talentos/Level (os dois
        podem ser clicados sem um esperar o outro).
        """
        t = I18N[self.lang]

        if getattr(self, "_sync_comparador_running", False):
            return  # já tem uma sincronização rodando, ignora clique duplicado

        active_digimons = getattr(self, "active_digimons", None)
        if not active_digimons:
            self.log(f" {t['msg_sync_excel_no_data']}", "alert")
            return

        self._sync_comparador_running = True
        self.btn_sync_comparador.config(state=tk.DISABLED, text=t["btn_sync_comparador_running"])
        self._cancel_sync_status_timer()
        self.lbl_sync_status.config(text="")

        def worker():
            result = self.sync_comparador_to_excel(active_digimons)
            self.root.after(0, lambda: self._on_sync_comparador_done(result))

        threading.Thread(target=worker, daemon=True).start()

    def _on_sync_comparador_done(self, result):
        """Roda na thread principal (chamado via root.after) pra atualizar a UI com segurança."""
        t = I18N[self.lang]
        self._sync_comparador_running = False
        self.btn_sync_comparador.config(state=tk.NORMAL, text=t["btn_sync_comparador"])

        status = result.get('status')
        label_msg, label_color = None, FG_COLOR

        if status == 'created':
            msg = t['msg_comparison_created'].format(count=result.get('count', 0))
            self.log(msg, "status")
            label_msg, label_color = msg.strip(), FG_COLOR
        elif status == 'updated':
            msg = t['msg_comparison_updated'].format(count=result.get('count', 0), new=result.get('new', 0))
            self.log(msg, "status")
            label_msg, label_color = msg.strip(), FG_COLOR
        elif status == 'reset':
            msg = t['msg_comparison_reset'].format(count=result.get('count', 0))
            self.log(msg, "status")
            label_msg, label_color = msg.strip(), FG_COLOR
        elif status == 'no_protected':
            msg = t['msg_sync_excel_none']
            self.log(f" {msg}", "status")
            label_msg, label_color = msg, FG_ALMOST
        elif status == 'disabled':
            msg = t['msg_comparador_disabled'] if not COMPARADOR_SYNC else t['msg_sync_excel_disabled']
            self.log(f" {msg}", "alert")
            label_msg, label_color = msg, FG_ALERT
        elif status == 'no_lib':
            msg = "⚠️ xlwings não instalado."
            self.log(" ⚠️ xlwings não instalado - sync com Excel desativado.", "alert")
            label_msg, label_color = msg, FG_ALERT
        elif status == 'error':
            comp_msg = f"{t['msg_comparison_error']}{result.get('error')}"
            self.log(comp_msg, "alert")
            label_msg, label_color = "⚠️ Erro na aba de comparação.", FG_ALERT

        if label_msg:
            self._show_sync_status_label(label_msg, label_color)

    def _cancel_sync_status_timer(self):
        if self._sync_status_after_id is not None:
            try:
                self.root.after_cancel(self._sync_status_after_id)
            except Exception:
                pass
            self._sync_status_after_id = None

    def _show_sync_status_label(self, message, color, duration_ms=3000):
        """Mostra 'message' no label logo abaixo dos botões de sync e agenda o sumiço
        automático depois de 'duration_ms'. Cancela qualquer timer pendente anterior
        pra evitar que um clique novo seja apagado por um timer de um clique antigo."""
        self._cancel_sync_status_timer()
        self.lbl_sync_status.config(text=message, fg=color)
        self._sync_status_after_id = self.root.after(duration_ms, self._clear_sync_status_label)

    def _clear_sync_status_label(self):
        self.lbl_sync_status.config(text="")
        self._sync_status_after_id = None

    def get_level_exp_display(self, digimon_id, level, current_exp):
        level_exp, next_level_exp = self.get_level_exp_progress(digimon_id, level, current_exp)
        if next_level_exp <= 0:
            return "MAX"
        return f"{level_exp}/{next_level_exp}"

    def get_level_exp_progress(self, digimon_id, level, current_exp):
        exp_base = get_exp_needed(digimon_id, level)
        exp_next = get_exp_needed(digimon_id, level + 1)
        if level >= 99 or exp_next <= exp_base:
            return 0, 0
        return max(0, current_exp - exp_base), exp_next - exp_base

    def build_progress_bar(self, remaining_xp, total_needed=None, length=10):
        max_remaining = MAX_LEVEL_BARRA
        remaining = max(0, min(max_remaining, remaining_xp))
        filled = int(((max_remaining - remaining) / max_remaining) * length)
        filled = max(0, min(length, filled))
        return ("█" * filled) + ("░" * (length - filled))

    def _compute_wishlist_matches(self, query_text):
        """Filtra self.active_digimons pelo texto digitado (ID exato ou substring do nome).
        Retorna (results_map, combo_options) — puro, sem tocar em nenhum widget."""
        t = I18N[self.lang]
        active = getattr(self, 'active_digimons', None) or {}

        is_id_search = query_text.isdigit()
        search_id = int(query_text) if is_id_search else None
        search_name = query_text.lower() if not is_id_search else None

        def match_criteria(d_id, d_name):
            if is_id_search:
                return d_id == search_id
            return search_name in d_name.lower()

        current_save = getattr(self, '_current_filename', '')
        results_map = []
        for info in active.values():
            if match_criteria(info['id'], info['name']):
                entry = dict(info)
                entry['save'] = current_save
                results_map.append(entry)

        if results_map:
            def sort_wishlist_result(info):
                level_exp, _ = self.get_level_exp_progress(info['id'], info['level'], info['exp'])
                return (info['level'], level_exp, info['talent'])
            results_map.sort(key=sort_wishlist_result, reverse=True)

        combo_options = []
        for info in results_map:
            exp_str = self.get_level_exp_display(info['id'], info['level'], info['exp'])
            lock_icon = "🔒 " if info.get('protected') else ""
            combo_options.append(f"{lock_icon}{info['name']} [{self._loc_display(info['loc'], info.get('slot'))}] - {t['lvl_abbr']}{info['level']}/{info['talent']} | EXP {exp_str} | Ref {info['uid']}")

        return results_map, combo_options

    def search_wishlist_digimon(self, silent=False):
        """Busca explícita (botão 'Buscar' ou Enter sem o dropdown aberto): preenche a
        combobox de resultado com TODOS os achados. Não usada mais pela digitação ao vivo
        (essa agora usa o dropdown flutuante — ver on_wishlist_query_keyrelease)."""
        t = I18N[self.lang]
        query_text = self.entry_wish_id.get().strip()
        if not query_text:
            if not silent:
                messagebox.showwarning(t["msg_warning_title"], t["wishlist_err_empty_query"])
            self.combo_wish_results['values'] = []
            self.combo_wish_results.set("")
            self.search_results_map = []
            self.lbl_wishlist_result_count.config(text="")
            return

        if not getattr(self, 'active_digimons', None):
            if not silent:
                messagebox.showwarning(t["msg_warning_title"], t["wishlist_err_no_save"])
            return

        self.search_results_map, combo_options = self._compute_wishlist_matches(query_text)

        if combo_options:
            self.combo_wish_results['values'] = combo_options
            self.combo_wish_results.current(0)
            result_word = "resultado" if len(combo_options) == 1 else "resultados"
            self.lbl_wishlist_result_count.config(text=f"{len(combo_options)} {result_word}")
        else:
            self.combo_wish_results['values'] = []
            self.combo_wish_results.set("")
            self.lbl_wishlist_result_count.config(text="0 resultados")
            if not silent:
                messagebox.showinfo(t["msg_search_title"], t["wishlist_no_results"].format(query=query_text))

    # ==========================================
    # AUTOCOMPLETE FLUTUANTE DA WISHLIST
    # Um dropdown próprio (não é a combobox nativa) que aparece embaixo do campo de busca
    # enquanto o usuário digita, sem NUNCA tirar o foco do campo — a combobox nativa do
    # Tkinter rouba o foco assim que o dropdown dela abre, o que quebrava a digitação.
    # ==========================================
    WISH_AC_MAX_ROWS = VALORES_BUSCA

    def on_wishlist_query_keyrelease(self, event=None):
        if event and event.keysym in ('Up', 'Down', 'Return', 'Escape', 'Tab'):
            return
        query_text = self.entry_wish_id.get().strip()
        if not query_text or not getattr(self, 'active_digimons', None):
            self.close_wish_autocomplete()
            return
        results_map, combo_options = self._compute_wishlist_matches(query_text)
        if not combo_options:
            self.close_wish_autocomplete()
            return
        self.show_wish_autocomplete(results_map[:self.WISH_AC_MAX_ROWS], combo_options[:self.WISH_AC_MAX_ROWS])

    def show_wish_autocomplete(self, results_map, combo_options):
        """Cria (se preciso) e preenche o dropdown flutuante, posicionado logo abaixo do
        campo de busca. Não chama focus_set em nada — o foco continua no Entry."""
        self._wish_ac_results = results_map
        self._wish_ac_options = combo_options
        self._wish_ac_index = 0

        if self._wish_ac_popup is None or not self._wish_ac_popup.winfo_exists():
            popup = tk.Toplevel(self.root)
            popup.overrideredirect(True)
            popup.attributes("-topmost", True)
            listbox = tk.Listbox(popup, bg="#1E1E1E", fg="white", font=("Consolas", 9),
                                  selectbackground="#1E90FF", selectforeground="white",
                                  activestyle="none", relief=tk.FLAT, highlightthickness=1,
                                  highlightbackground="#444444", bd=0, exportselection=False)
            listbox.pack(fill=tk.BOTH, expand=True)
            listbox.bind("<ButtonRelease-1>", self.on_wish_autocomplete_click)
            self._wish_ac_popup = popup
            self._wish_ac_listbox = listbox

        x = self.entry_wish_id.winfo_rootx()
        y = self.entry_wish_id.winfo_rooty() + self.entry_wish_id.winfo_height()
        width = self.entry_wish_id.winfo_width()
        row_h = 18
        height = row_h * min(len(combo_options), self.WISH_AC_MAX_ROWS) + 6
        self._wish_ac_popup.geometry(f"{width}x{height}+{x}+{y}")

        self._wish_ac_listbox.delete(0, tk.END)
        for option in combo_options:
            self._wish_ac_listbox.insert(tk.END, f" {option}")
        self._wish_ac_listbox.selection_clear(0, tk.END)
        self._wish_ac_listbox.selection_set(0)
        self._wish_ac_popup.deiconify()

    def close_wish_autocomplete(self):
        if getattr(self, '_wish_ac_popup', None) is not None and self._wish_ac_popup.winfo_exists():
            self._wish_ac_popup.withdraw()
        self._wish_ac_index = -1

    def on_wishlist_query_down(self, event=None):
        """Seta pra baixo: move o destaque dentro do dropdown flutuante SEM tirar o foco do campo."""
        if self._wish_ac_popup is None or str(self._wish_ac_popup.state()) != "normal" or not getattr(self, '_wish_ac_options', None):
            self.on_wishlist_query_keyrelease()  # nada aberto ainda: tenta abrir com o texto atual
            return "break"
        self._wish_ac_index = min(self._wish_ac_index + 1, len(self._wish_ac_options) - 1)
        self._wish_ac_listbox.selection_clear(0, tk.END)
        self._wish_ac_listbox.selection_set(self._wish_ac_index)
        self._wish_ac_listbox.see(self._wish_ac_index)
        return "break"

    def on_wishlist_query_up(self, event=None):
        if self._wish_ac_popup is None or str(self._wish_ac_popup.state()) != "normal" or not getattr(self, '_wish_ac_options', None):
            return "break"
        self._wish_ac_index = max(self._wish_ac_index - 1, 0)
        self._wish_ac_listbox.selection_clear(0, tk.END)
        self._wish_ac_listbox.selection_set(self._wish_ac_index)
        self._wish_ac_listbox.see(self._wish_ac_index)
        return "break"

    def on_wishlist_query_return(self, event=None):
        """Enter: se o dropdown flutuante estiver aberto com algo destacado, confirma essa
        escolha. Senão, cai pro comportamento antigo (busca explícita)."""
        if self._wish_ac_popup is not None and str(self._wish_ac_popup.state()) == "normal" and getattr(self, '_wish_ac_options', None):
            self.select_wish_autocomplete_index(self._wish_ac_index)
        else:
            self.search_wishlist_digimon()
    
    def on_wishlist_tab_to_combo(self, event):
        # 1. Verifica se o dropdown flutuante está aberto e tem opções
        if self._wish_ac_popup is not None and str(self._wish_ac_popup.state()) == "normal" and getattr(self, '_wish_ac_options', None):
            # Confirma a sugestão destacada (joga pra combobox e pro campo de busca)
            self.select_wish_autocomplete_index(self._wish_ac_index)
            
            # Pulo Ninja: Já manda o cursor direto para digitar o Nível Alvo!
            self.entry_wish_target.focus_set()
            return "break"
        
        # 2. Se o dropdown estiver fechado, o TAB faz o comportamento normal de pular
        self.combo_wish_results.focus_set()
        
        # Retorna "break" para cancelar a ação nativa do Tkinter
        return "break"

    def on_wishlist_query_escape(self, event=None):
        self.close_wish_autocomplete()

    def on_wish_autocomplete_click(self, event=None):
        index = self._wish_ac_listbox.nearest(event.y)
        self.select_wish_autocomplete_index(index)

    def select_wish_autocomplete_index(self, index):
        """Confirma a sugestão destacada: preenche a combobox de resultado (é o que
        add_wishlist_target usa) e fecha o dropdown flutuante, sem nunca ter tirado
        o foco do campo de texto durante a digitação."""
        options = getattr(self, '_wish_ac_options', [])
        results = getattr(self, '_wish_ac_results', [])
        if not options or not (0 <= index < len(options)):
            return
        self.search_results_map = results
        self.combo_wish_results['values'] = options
        self.combo_wish_results.current(index)
        result_word = "resultado" if len(options) == 1 else "resultados"
        self.lbl_wishlist_result_count.config(text=f"{len(options)} {result_word}")
        self.entry_wish_id.delete(0, tk.END)
        self.entry_wish_id.insert(0, results[index]['name'])
        self.close_wish_autocomplete()
        self.entry_wish_id.focus_set()
        self.entry_wish_id.icursor(tk.END)

    def add_wishlist_target(self):
        """Adiciona o Digimon selecionado à Wishlist com a meta de nível e salva no JSON."""
        t = I18N[self.lang]
        idx = self.combo_wish_results.current()
        if idx < 0 or not hasattr(self, 'search_results_map') or idx >= len(self.search_results_map):
            messagebox.showwarning(t["msg_warning_title"], t["wishlist_err_select_first"])
            return

        target_str = self.entry_wish_target.get().strip()
        if not target_str.isdigit():
            messagebox.showwarning(t["msg_warning_title"], t["wishlist_err_invalid_target"])
            return

        target_lvl = int(target_str)
        if not (1 <= target_lvl <= 99):
            messagebox.showwarning(t["msg_warning_title"], t["wishlist_err_target_range"])
            return

        selected_info = self.search_results_map[idx]

        # REGRA: o Target Level não pode ser MAIOR que o Talento/Limite atual do Digimon
        # (acima do talento ele evoluiria antes de chegar lá, e essa meta nunca seria alcançada)
        if target_lvl > selected_info['talent']:
            messagebox.showwarning(t["msg_warning_title"], t["wishlist_err_target_too_high"].format(tgt=target_lvl, cap=selected_info['talent']))
            return

        item = {
            'uid': selected_info['uid'],
            'id': selected_info['id'],
            'name': selected_info['name'],
            'loc': selected_info['loc'],
            'slot': selected_info['slot'],
            'target_lvl': target_lvl,
            'orphaned': False,
            'save': selected_info.get('save', '')
        }
        self.wishlist.append(item)
        self.save_config() 

        self.combo_wish_results.set("")
        self.combo_wish_results['values'] = []
        self.lbl_wishlist_result_count.config(text="")
        self.entry_wish_target.delete(0, tk.END)
        self.entry_wish_id.delete(0, tk.END)

        if hasattr(self, '_current_filepath') and os.path.exists(self._current_filepath):
            self.process_save(self._current_filepath, self._current_filename)
            
        # ---> MUDANÇA AQUI <---
        # Limpou tudo? Joga o foco de volta na busca inicial!
        self.entry_wish_id.focus_set()

    def delete_wishlist_item(self, original_index):
        """Remove um item da Wishlist pelo índice (pra sempre) e atualiza o JSON."""
        if 0 <= original_index < len(self.wishlist):
            del self.wishlist[original_index]
            self.save_config()  # Salva a alteração no config.json
            if hasattr(self, '_current_filepath') and os.path.exists(self._current_filepath):
                self.process_save(self._current_filepath, self._current_filename)

    def readd_wishlist_item(self, original_index):
        """Tenta reativar um item órfão da Wishlist, checando se o Digimon está de volta no save atual."""
        t = I18N[self.lang]
        if not (0 <= original_index < len(self.wishlist)):
            return

        item = self.wishlist[original_index]
        active = getattr(self, 'active_digimons', {})

        found = None
        item_uid = item.get('uid')
        if item_uid:
            for candidate in active.values():
                if candidate.get('uid') == item_uid and candidate['id'] == item['id'] and candidate['name'] == item['name']:
                    found = candidate
                    break
        else:
            key = (item.get('loc'), item.get('slot'))
            if key in active and active[key]['id'] == item['id'] and active[key]['name'] == item['name']:
                found = active[key]
            else:
                candidates = [candidate for candidate in active.values() if candidate['id'] == item['id'] and candidate['name'] == item['name']]
                if len(candidates) == 1:
                    found = candidates[0]

        if found:
            item['orphaned'] = False
            item['uid'] = found.get('uid', item.get('uid', ''))
            item['loc'] = found['loc']
            item['slot'] = found['slot']
            self.save_config()
            if hasattr(self, '_current_filepath') and os.path.exists(self._current_filepath):
                self.process_save(self._current_filepath, self._current_filename)
        else:
            messagebox.showwarning(t["msg_warning_title"], t["wishlist_not_found_msg"].format(name=item['name']))

    def update_combo_list(self, event=None):
        available_saves = []
        if self.save_dir and os.path.exists(self.save_dir):
            for i in range(16):
                filename = f"{i:04d}.bin"
                if os.path.exists(os.path.join(self.save_dir, filename)):
                    available_saves.append(filename)
        self.save_combo['values'] = available_saves

    def get_most_recent_save(self):
        """Save modificado mais recentemente entre todos os que existem (incluindo o auto-save 0000.bin,
        que é o mais usado na prática já que o jogo escreve nele o tempo todo)."""
        if not self.save_dir or not os.path.exists(self.save_dir):
            return None
        latest_name, max_mtime = None, 0
        for i in range(16):
            fname = f"{i:04d}.bin"
            fpath = os.path.join(self.save_dir, fname)
            if os.path.exists(fpath):
                mtime = os.path.getmtime(fpath)
                if mtime > max_mtime:
                    max_mtime = mtime
                    latest_name = fname
        return latest_name

    def update_wishlist_context_label(self):
        """Atualiza o aviso contextual da Wishlist: só aparece quando pausado/inspecionando um save
        específico, avisando que dá pra adicionar Digimons desse save na Wishlist ali mesmo."""
        t = I18N[self.lang]
        filename = getattr(self, '_current_filename', None)
        if self.is_paused and filename:
            self.lbl_wishlist_context.config(text=t["wishlist_ctx_paused"].format(save=filename))
            self.lbl_wishlist_context.pack(fill='x', padx=15, pady=(0, 8), before=self.lbl_wishlist_search)
        else:
            self.lbl_wishlist_context.pack_forget()

    def log(self, text, tag=None, widget=None):
        target = widget if widget is not None else self.text_area
        target.config(state=tk.NORMAL)
        if isinstance(text, list):
            for pedaco, t in text:
                target.insert(tk.END, pedaco, t)
            target.insert(tk.END, "\n")
        else:
            target.insert(tk.END, text + "\n", tag)
        target.config(state=tk.DISABLED)
        target.yview(1.0)

    def log_lines(self, entries, widget=None):
        """Insere várias linhas de uma vez, ligando/desligando o estado do widget e rolando a
        tela só UMA vez no total — em vez de uma vez por linha (self.log). Essencial pra listas
        grandes (centenas/milhares de linhas), onde o custo por linha do self.log() normal vira
        o gargalo de performance. Cada item de 'entries' segue o mesmo formato aceito por self.log:
        uma string, ou uma lista de tuplas (texto, tag) pra múltiplas cores na mesma linha."""
        target = widget if widget is not None else self.text_area
        target.config(state=tk.NORMAL)
        for text, tag in entries:
            if isinstance(text, list):
                for pedaco, seg_tag in text:
                    target.insert(tk.END, pedaco, seg_tag)
                target.insert(tk.END, "\n")
            else:
                target.insert(tk.END, text + "\n", tag)
        target.config(state=tk.DISABLED)
        target.yview(1.0)

    def update_summary_panel(self, count_party, count_box, count_farm, total_alerts, total_almost, total_lv99, total_protected, t, loc_lbl):
        """Preenche o painel de Resumo na barra lateral: ocupação de Party/Box/Fazenda (labels grandes) + resumo de pendências (texto)."""
        self.lbl_party_count.config(text=f"{loc_lbl['PARTY']}: {count_party}/6")
        self.lbl_box_count.config(text=f"{loc_lbl['BOX']}: {count_box}/999")
        self.lbl_farm_count.config(text=f"{loc_lbl['FAZENDA']}: {count_farm}/30")
        self.lbl_total_count.config(text=f"TOTAL: {count_party + count_box + count_farm}")

        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.config(state=tk.DISABLED)

        if total_alerts == 0 and total_almost == 0 and total_lv99 == 0 and total_protected == 0:
            self.log(t["msg_all_normal"], "status", widget=self.summary_text)
        else:
            if total_alerts > 0:
                self.log(f" -> {total_alerts}{t['summary_alerts']}", "alert", widget=self.summary_text)
            if total_almost > 0:
                self.log(f" -> {total_almost}{t['summary_almost']}", "almost", widget=self.summary_text)
            if total_lv99 > 0:
                self.log(f" -> {total_lv99}{t['summary_lv99']}", "loc_fazenda", widget=self.summary_text)
            if total_protected > 0:
                self.log(f" -> {total_protected}{t['summary_protected']}", "status", widget=self.summary_text)

    def get_quase_button_label(self, t):
        if self.quase_list_max_count == 0:
            return t["btn_quase_show_until_all"]
        if self.quase_list_max_count == QUASE_LIST_CYCLE[0]:
            return t["btn_quase_show_until_10"]
        if self.quase_list_max_count == QUASE_LIST_CYCLE[1]:
            return t["btn_quase_show_until_30"]
        return t["btn_quase_show_until_custom"].format(count=self.quase_list_max_count)

    def cycle_quase_max_count(self, filepath, filename):
        current = self.quase_list_max_count
        cycle = QUASE_LIST_CYCLE
        if current in cycle:
            idx = cycle.index(current)
            next_idx = (idx + 1) % len(cycle)
            self.quase_list_max_count = cycle[next_idx]
        else:
            self.quase_list_max_count = cycle[0]
        self.quase_limit_entry_var.set("")
        self.save_config()
        self.process_save(filepath, filename)

    def apply_quase_limit_from_entry(self, filepath, filename, event=None):
        value = self.quase_limit_entry_var.get().strip()
        if value == "":
            self.quase_list_max_count = 0
        elif value.isdigit():
            count = int(value)
            self.quase_list_max_count = 0 if count <= 0 else count
        else:
            messagebox.showwarning(I18N[self.lang]["msg_warning_title"], I18N[self.lang]["msg_quase_invalid_count"])
            return
        self.save_config()
        self.process_save(filepath, filename)

    def apply_quase_filters(self, filepath, filename):
        self.quase_filter_location = self.quase_filter_loc_var.get()
        self.quase_filter_protected = self.quase_filter_protected_var.get()
        self.quase_filter_wishlist = self.quase_filter_wishlist_var.get()
        self.save_config()
        self.process_save(filepath, filename)

    def on_quase_search_keyrelease(self, filepath, filename, event=None):
        """Busca ao vivo dentro da lista Quase Lá. Como reaplicar o filtro exige reprocessar
        o save inteiro (openssl + reescanear tudo), usamos um pequeno atraso ('debounce') pra
        só disparar depois que o usuário parar de digitar por um instante — assim a digitação
        continua fluida em vez de travar a cada tecla."""
        if event and event.keysym in ('Up', 'Down', 'Return', 'Escape', 'Tab'):
            return
        if self._quase_search_after_id:
            self.root.after_cancel(self._quase_search_after_id)
        self._quase_search_after_id = self.root.after(350, lambda: self._commit_quase_search(filepath, filename))

    def _commit_quase_search(self, filepath, filename):
        self._quase_search_after_id = None
        self.quase_search_text = self.quase_search_var.get()
        if hasattr(self, '_current_filepath') and os.path.exists(self._current_filepath):
            self.process_save(self._current_filepath, self._current_filename)
        # process_save reconstrói a tela inteira (inclusive esse campo) — devolve o foco e o
        # cursor pro fim do texto, senão o usuário perderia o foco a cada pausa na digitação.
        if hasattr(self, 'entry_quase_search') and self.entry_quase_search.winfo_exists():
            self.entry_quase_search.focus_set()
            self.entry_quase_search.icursor(tk.END)

    def clear_quase_search(self, filepath, filename):
        """Enter no campo de busca da Quase Lá: limpa a busca por completo e volta a lista
        ao tamanho normal, respeitando os outros filtros que continuarem ativos."""
        if self._quase_search_after_id:
            self.root.after_cancel(self._quase_search_after_id)
            self._quase_search_after_id = None
        self.quase_search_var.set("")
        self.quase_search_text = ""
        if hasattr(self, '_current_filepath') and os.path.exists(self._current_filepath):
            self.process_save(self._current_filepath, self._current_filename)

    def on_mode_change(self):
        self.is_paused = False
        self.lbl_paused.pack_forget()
        self.btn_resume.pack_forget()
        self.last_mtime = 0
        self._pending_save_check = None
        self.update_combo_list()
        self.root.focus_set()

    def on_combo_select(self, event):
        self.is_paused = True
        self.lbl_paused.pack(pady=5)
        self.btn_resume.pack(pady=10, fill=tk.X)
        self.save_combo.selection_clear()
        self.root.focus_set()
        
        selected_file = self.save_combo.get()
        filepath = os.path.join(self.save_dir, selected_file)
        if os.path.exists(filepath):
            self.process_save(filepath, selected_file)

    def resume_tracking(self):
        self.is_paused = False
        self.lbl_paused.pack_forget()
        self.btn_resume.pack_forget()
        self.last_mtime = 0
        self._pending_save_check = None
        self.btn_resume.config(bg="#8B0000") 
        self.save_combo.selection_clear()
        self.root.focus_set()
        self.update_wishlist_context_label()

    def get_target_save(self):
        if not self.save_dir or not os.path.exists(self.save_dir):
            return None, 0
            
        latest_file = None
        max_mtime = 0
        start_idx = 0 if self.mode.get() == "AUTO" else 1
        
        for i in range(start_idx, 16):
            filepath = os.path.join(self.save_dir, f"{i:04d}.bin")
            if os.path.exists(filepath):
                mtime = os.path.getmtime(filepath)
                if mtime > max_mtime:
                    max_mtime = mtime
                    latest_file = filepath
                    
        return latest_file, max_mtime
        
    def _loc_display(self, loc, slot=None, translate=True):
        """
        Retorna o texto usado dentro dos colchetes de localização (ex.: "[ PARTY 3 ]").

        - PARTY: inclui a posição do Digimon na equipe (1-6) -> "PARTY 1".."PARTY 6".
          Isso sempre dá 7 caracteres ("PARTY " + 1 dígito), o MESMO tamanho de
          "FAZENDA" - então o alinhamento das listas (que usa :^7) continua
          idêntico, sem precisar mudar nenhum width.
        - BOX / FAZENDA: comportamento igual a antes. translate=True usa a
          tradução (loc_labels) pro texto em tela; translate=False devolve o
          valor cru ("BOX"/"FAZENDA"), usado na aba do Excel (que não deve
          variar com o idioma da interface).
        - slot é 0-based (mesmo índice salvo em active_digimons/'slot'),
          por isso o +1 pra virar "PARTY 1" em vez de "PARTY 0".
        """
        if loc == "PARTY" and slot is not None:
            return f"PARTY {slot + 1}"
        if translate:
            return I18N[self.lang]["loc_labels"].get(loc, loc)
        return loc

    def calculate_almost_data(self, data, name_offset, name, level, talent, loc, protected, uid, is_almost, slot=None):
        digimon_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
        current_exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
        ascendant_talent_raw = struct.unpack_from("<I", data, name_offset + 0xFC)[0]
        
        
        exp_alvo = get_exp_needed(digimon_id, talent)
        exp_base = get_exp_needed(digimon_id, level)
        
        faltam = exp_alvo - current_exp
        
        progresso_total = exp_alvo - exp_base
        bar_str = self.build_progress_bar(faltam, progresso_total)
        faltam_str = f"{faltam:,}".replace(",", ".")
        
        return (name, level, talent, faltam, bar_str, faltam_str, loc, protected, uid, is_almost, ascendant_talent_raw, slot)

    def process_save(self, filepath, filename):
        self._current_filepath = filepath
        self._current_filename = filename
        self.active_digimons = {}
        OPENSSL_PATH = os.path.join(SCRIPT_DIR, "openssl.exe")
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete(1.0, tk.END)
        
        t = I18N[self.lang]
        loc_lbl = t["loc_labels"]
        if not os.path.exists(OPENSSL_PATH):
            self.log("=" * 75, "header")
            self.log(t["err_critical_title"], "alert")
            self.log("=" * 75, "header")
            self.log(t["err_openssl_notfound"], "almost")
            self.log(f"{t['err_openssl_path']} {SCRIPT_DIR}\n", "status")
            return

        try:
            cmd = [OPENSSL_PATH, "enc", "-d", "-aes-128-ecb", "-K", AES_KEY, "-in", filepath, "-out", SAVE_FILE_DEC, "-nopad"]
            CREATE_NO_WINDOW = 0x08000000
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, creationflags=CREATE_NO_WINDOW)
        except Exception as e:
            self.log(f"{t['err_decrypt']} {e}", "alert")
            return

        if not os.path.exists(SAVE_FILE_DEC): 
            return

        with open(SAVE_FILE_DEC, "rb") as f:
            data = f.read()

        alerts = {"PARTY": [], "BOX": [], "FAZENDA": []}
        lv99_list = {"PARTY": [], "BOX": [], "FAZENDA": []} 
        almost_list = [] 
        remaining_list = []
        total_alerts, total_almost, total_lv99, total_protected = 0, 0, 0, 0 
        
        count_party = 0
        count_box = 0
        count_farm = 0
        
        processed = set()
        active_digimons = {}

        # Configuração exata da Fazenda
        FARM_START = 0x539C8
        FARM_SIZE = 344
        HEADER_FARM = 0x18

        for i in range(30):
            offset = FARM_START + (i * FARM_SIZE)
            if offset + FARM_SIZE > len(data) or offset in processed: 
                continue

            if data[offset] == 1:
                name_offset = offset + HEADER_FARM
                
                name_bytes = bytearray()
                for b in data[name_offset : name_offset + 32]:
                    if b == 0: break
                    name_bytes.append(b)
                    
                try:
                    name = name_bytes.decode('ascii')
                    if len(name) < 2 or not (65 <= ord(name[0]) <= 90): 
                        continue
                except: 
                    continue

                count_farm += 1
                processed.add(offset)

                # Dentro de for i in range(30):
                digimon_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
                uid = self.read_digimon_uid(data, name_offset)
                level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                current_exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                ascendant_talent_raw = struct.unpack_from("<I", data, name_offset + 0xFC)[0]
                talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]
                elo_raw = struct.unpack_from("<f", data, name_offset + 0x13C)[0]
                elo = int(elo_raw / 100)
                protected = self.read_digimon_protected(data, name_offset)
                talent_for_dict = min(talent_raw // 1000, 99) if talent_raw >= 1000 else 1

                active_digimons[("FAZENDA", i)] = {
                    'uid': uid,
                    'id': digimon_id,
                    'name': name,
                    'level': level,
                    'exp': current_exp,
                    'loc': "FAZENDA",
                    'slot': i,
                    'protected': protected,
                    'talent': talent_for_dict,
                    'ascendant_talent': ascendant_talent_raw,
                    'elo': elo
                }

                if talent_raw >= 1000:
                    talent = talent_raw // 1000
                    if talent > 99: talent = 99 
                    
                    if level == 99:
                        lv99_list["FAZENDA"].append((name, level, talent, protected, ascendant_talent_raw, i))
                        total_lv99 += 1
                    elif level >= talent: 
                        alerts["FAZENDA"].append((name, level, talent, protected, ascendant_talent_raw, i))
                        total_alerts += 1
                    elif level == talent - 1:
                        almost_data = self.calculate_almost_data(data, name_offset, name, level, talent, "FAZENDA", protected, uid, True, i)
                        almost_list.append(almost_data)
                        total_almost += 1
                    elif level < talent - 1:
                        remaining_data = self.calculate_almost_data(data, name_offset, name, level, talent, "FAZENDA", protected, uid, False, i)
                        if remaining_data[3] > 0:
                            remaining_list.append(remaining_data)

                    if protected:
                        total_protected += 1

        regions = [("PARTY", 0x12C8 - 0x10, 0x10, 6), ("BOX", 0x1AA8 - 0x10, 0x10, 999)]
        for loc, start, h_size, max_s in regions:
            for i in range(max_s):
                offset = start + (i * DIGIMON_SIZE)
                if offset + DIGIMON_SIZE > len(data) or offset in processed or data[offset] != 1: continue
                
                name_offset = offset + h_size
                name_bytes = bytearray()
                for b in data[name_offset : name_offset + 32]:
                    if b == 0: break
                    name_bytes.append(b)
                    
                try:
                    name = name_bytes.decode('ascii')
                    if len(name) < 2 or not (65 <= ord(name[0]) <= 90): continue
                except: continue

                if loc == "PARTY":
                    count_party += 1
                elif loc == "BOX": 
                    count_box += 1
                
                processed.add(offset)

                # Dentro de for loc, start, h_size, max_s in regions: ... for i in range(max_s):
                digimon_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
                uid = self.read_digimon_uid(data, name_offset)
                level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                current_exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                ascendant_talent_raw = struct.unpack_from("<I", data, name_offset + 0xFC)[0]
                talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]
                elo_raw = struct.unpack_from("<f", data, name_offset + 0x13C)[0]
                elo = int(elo_raw / 100)
                protected = self.read_digimon_protected(data, name_offset)
                talent_for_dict = min(talent_raw // 1000, 99) if talent_raw >= 1000 else 1

                active_digimons[(loc, i)] = {
                    'uid': uid,
                    'id': digimon_id,
                    'name': name,
                    'level': level,
                    'exp': current_exp,
                    'loc': loc,
                    'slot': i,
                    'protected': protected,
                    'talent': talent_for_dict,
                    'ascendant_talent': ascendant_talent_raw,
                    'elo': elo
                }

                if talent_raw >= 1000:
                    talent = talent_raw // 1000
                    if talent > 99: talent = 99 
                    
                    if level == 99:
                        lv99_list[loc].append((name, level, talent, protected, ascendant_talent_raw, i))
                        total_lv99 += 1
                    elif level >= talent: 
                        alerts[loc].append((name, level, talent, protected, ascendant_talent_raw, i))
                        total_alerts += 1
                    elif level == talent - 1:
                        almost_data = self.calculate_almost_data(data, name_offset, name, level, talent, loc, protected, uid, True, i)
                        almost_list.append(almost_data)
                        total_almost += 1
                    elif level < talent - 1:
                        remaining_data = self.calculate_almost_data(data, name_offset, name, level, talent, loc, protected, uid, False, i)
                        if remaining_data[3] > 0:
                            remaining_list.append(remaining_data)

                    if protected:
                        total_protected += 1

        # ==========================================
        # RESOLVE WISHLIST CONTRA O SAVE ATUAL
        # (precisa rodar antes da 1ª lista, pra podermos "transferir"
        #  quem já bateu a meta pra lá)
        #
        # Itens que sumirem do save NÃO são apagados na hora -> ficam
        # marcados como "órfãos" e aparecem numa notificação com botões
        # de Readicionar / Remover. Só saem de vez quando o usuário confirma.
        # ==========================================
        self.active_digimons = active_digimons  # guarda pra os botões de Readicionar usarem depois
        self.update_wishlist_context_label()

        # OBS: a sincronização com o Excel NÃO roda mais automaticamente aqui.
        # Ela ficava cara demais rodando a cada leitura de save (que acontece
        # com bastante frequência no modo automático). Agora é sob demanda,
        # via dois botões independentes: "🔄 Sync Main Sheet" (self.btn_sync_talentos
        # -> self.sync_protected_talents_to_excel) e "📊 Sync Comparator"
        # (self.btn_sync_comparador -> self.sync_comparador_to_excel).

        wishlist_resolved = []   # [(w_idx, item, dig_info), ...] -> ainda existem no save, ativos
        wishlist_orphaned = []   # [(w_idx, item), ...] -> órfãos (já marcados, ou detectados agora)
        wishlist_state_changed = False

        for w_idx, item in enumerate(self.wishlist):
            if item.get('orphaned'):
                wishlist_orphaned.append((w_idx, item))
                continue

            dig_info = None
            item_uid = item.get('uid')

            if item_uid:
                for candidate in active_digimons.values():
                    if candidate.get('uid') == item_uid and candidate['id'] == item['id'] and candidate['name'] == item['name']:
                        dig_info = candidate
                        break
            else:
                key = (item.get('loc'), item.get('slot'))
                if key in active_digimons and active_digimons[key]['id'] == item['id'] and active_digimons[key]['name'] == item['name']:
                    dig_info = active_digimons[key]
                else:
                    candidates = [candidate for candidate in active_digimons.values() if candidate['id'] == item['id'] and candidate['name'] == item['name']]
                    if len(candidates) == 1:
                        dig_info = candidates[0]

            if dig_info:
                # Guarda de qual save esse Digimon foi visto pela última vez (só informativo, não trava mais nada)
                if item.get('save') != filename:
                    item['save'] = filename
                    wishlist_state_changed = True
                if item.get('uid') != dig_info.get('uid'):
                    item['uid'] = dig_info.get('uid', '')
                    wishlist_state_changed = True
                if item.get('loc') != dig_info['loc'] or item.get('slot') != dig_info['slot']:
                    item['loc'] = dig_info['loc']
                    item['slot'] = dig_info['slot']
                    wishlist_state_changed = True
                wishlist_resolved.append((w_idx, item, dig_info))
            elif not self.is_paused:
                # Digimon sumiu do save que está sendo monitorado ao vivo (evoluiu, foi deletado,
                # ou o save foi sobrescrito) -> vira órfão. Só faz isso durante monitoramento ao vivo —
                # nunca enquanto o usuário está só inspecionando/dando uma olhada num save (is_paused=True).
                item['orphaned'] = True
                wishlist_orphaned.append((w_idx, item))
                wishlist_state_changed = True
            # else: pausado/inspecionando e não achou -> ignora silenciosamente neste ciclo, sem marcar nada.

        if wishlist_state_changed:
            self.save_config()

        # Calcula o progresso de cada item ativo e separa quem já bateu a meta
        wishlist_pending = []   # ainda não bateu -> aparece na 4ª lista normalmente
        wishlist_reached = []   # bateu a meta -> "transferido" pra 1ª lista (Limit Cap)

        for w_idx, item, dig_info in wishlist_resolved:
            dig_id = item['id']
            target_lvl = item['target_lvl']
            cur_lvl = dig_info['level']
            cur_exp = dig_info['exp']
            w_loc = dig_info['loc']
            w_name = dig_info['name']
            w_protected = dig_info.get('protected', False)
            w_slot = dig_info.get('slot')

            # Dentro do loop de wishlist_resolved:
            exp_target = get_exp_needed(dig_id, target_lvl)
            exp_base = get_exp_needed(dig_id, cur_lvl)
            faltam = exp_target - cur_exp
            asc_talent = dig_info['ascendant_talent']

            if faltam <= 0 or cur_lvl >= target_lvl:
                wishlist_reached.append((w_idx, w_name, cur_lvl, target_lvl, w_loc, w_protected, asc_talent, w_slot))
            else:
                prog_total = exp_target - exp_base
                bar_str = self.build_progress_bar(faltam, prog_total)
                faltam_str = f"{faltam:,}".replace(",", ".")
                wishlist_pending.append((faltam, w_idx, w_name, cur_lvl, target_lvl, bar_str, faltam_str, w_loc, w_protected, asc_talent, w_slot))

        # Os que bateram a meta contam no resumo junto com os alertas normais
        total_alerts += len(wishlist_reached)

        self.save_combo.set(filename)
        self.save_combo.selection_clear()
        
        # Onde a mensagem de Status é impressa:
        self.log("=" * 75, "header")
        self.log(t["app_title"], "header")
        self.log("=" * 75, "header")
        
        
        status_msg = f"{t['status_inspecting']}{filename}" if self.is_paused else f"{t['status_monitoring']}{filename}"
        self.log(f"{t['status_prefix']}{status_msg}{t['lbl_updated']}{time.strftime('%H:%M:%S')}", "status")
        self.log(t.get("lbl_obs_acc", "OBS: Os valores entre chaves{} nas linhas representam o novo valor do talento se você evoluir o Digimon."), "almost")
        self.log("-" * 75, "status")

        if wishlist_orphaned:
            self.log(t.get("wishlist_auto_removed", "⚠️  WISHLIST: digimon(s) não encontrado(s) no save atual:"), "alert")
            for w_idx, orphan_item in wishlist_orphaned:
                self.text_area.config(state=tk.NORMAL)
                self.text_area.insert(tk.END, f"    - {orphan_item['name']:<14} (Target Lv. {orphan_item['target_lvl']:02d})  ", "almost")

                btn_readd = tk.Button(self.text_area, text=t["btn_wishlist_readd"], command=lambda idx=w_idx: self.readd_wishlist_item(idx),
                                    bg="#444444", fg="#1E90FF", font=("Consolas", 8, "bold"),
                                    relief=tk.FLAT, cursor="hand2", padx=4, pady=0)
                self.text_area.window_create(tk.END, window=btn_readd)
                self.text_area.insert(tk.END, "  ")

                btn_forget = tk.Button(self.text_area, text=t["btn_wishlist_forget"], command=lambda idx=w_idx: self.delete_wishlist_item(idx),
                                    bg="#444444", fg="red", font=("Consolas", 8, "bold"),
                                    relief=tk.FLAT, cursor="hand2", padx=4, pady=0)
                self.text_area.window_create(tk.END, window=btn_forget)

                self.text_area.insert(tk.END, "\n")
                self.text_area.config(state=tk.DISABLED)
            self.log("-" * 75, "status")

        txt_atingiu = I18N[self.lang].get("atingiu", "<-- ATINGIU O LIMITE!")
        txt_quase = I18N[self.lang].get("quase", " ⏳ QUASE LÁ (Faltando 1 Level):")

        reached_rows = []
        for loc in ["PARTY", "BOX", "FAZENDA"]:
            alerts[loc].sort(key=lambda x: (x[0], x[1]))
            # 1. ADICIONADO: asc_talent + slot no desempacotamento
            for name, level, talent, protected, asc_talent, slot in alerts[loc]:
                reached_rows.append(("alert", loc, name, level, talent, protected, asc_talent, slot))

            # 2. ADICIONADO: asc_talent + slot no desempacotamento da wishlist
            for w_idx, w_name, w_level, w_target, w_item_loc, w_protected, asc_talent, w_slot in wishlist_reached:
                if w_item_loc != loc:
                    continue
                reached_rows.append(("wishlist", loc, w_idx, w_name, w_level, w_target, w_protected, asc_talent, w_slot))

        self.update_summary_panel(count_party, count_box, count_farm, total_alerts, total_almost, total_lv99, total_protected, t, loc_lbl)

        # ==========================================
        # LISTAS 2 a 4 — ORDEM CUSTOMIZÁVEL PELO USUÁRIO (setas ▲▼ no cabeçalho)
        # A ordem escolhida fica salva em self.list_order / config.json
        # ==========================================

        def render_header_with_controls(label_text, tag, show_attr, btn_show_key, list_key, is_first, is_last):
            """Cabeçalho de uma lista expansível: rótulo + setas de reordenar + botão mostrar/ocultar, tudo na mesma linha."""
            def toggle():
                setattr(self, show_attr, not getattr(self, show_attr, False))
                self.save_config()
                self.process_save(filepath, filename)

            btn_text = t["btn_hide_details"] if getattr(self, show_attr, False) else t[btn_show_key]

            self.text_area.config(state=tk.NORMAL)
            self.text_area.insert(tk.END, label_text, tag)

            controls_width = len(btn_text) + 10  # reserva espaço extra pras setas ▲▼
            padding = max(2, 75 - len(label_text) - controls_width)
            self.text_area.insert(tk.END, " " * padding, tag)

            if not is_first:
                btn_up = tk.Button(self.text_area, text="▲", command=lambda: self.move_list_order(list_key, -1),
                                    bg="#333333", fg="white", font=("Consolas", 8, "bold"),
                                    relief=tk.FLAT, cursor="hand2", padx=3, pady=0)
                self.text_area.window_create(tk.END, window=btn_up)
            if not is_last:
                btn_down = tk.Button(self.text_area, text="▼", command=lambda: self.move_list_order(list_key, 1),
                                      bg="#333333", fg="white", font=("Consolas", 8, "bold"),
                                      relief=tk.FLAT, cursor="hand2", padx=3, pady=0)
                self.text_area.window_create(tk.END, window=btn_down)

            btn_toggle = tk.Button(self.text_area, text=btn_text, command=toggle,
                            bg="#333333", fg="white", font=("Consolas", 8, "bold"),
                            relief=tk.FLAT, cursor="hand2", padx=8, pady=0)
            self.text_area.window_create(tk.END, window=btn_toggle)
            self.text_area.insert(tk.END, "\n\n")
            self.text_area.config(state=tk.DISABLED)

        def append_section_divider():
            self.text_area.config(state=tk.NORMAL)
            self.text_area.insert(tk.END, "-" * 75 + "\n", "status")
            self.text_area.config(state=tk.DISABLED)

        def render_almost_list(is_first, is_last):
            total_quase = len(almost_list) + len(remaining_list)
            if total_quase <= 0:
                return
            render_header_with_controls(txt_quase, "header_yellow", 'show_almost', 'btn_show_almost', 'almost', is_first, is_last)

            if getattr(self, 'show_almost', False):
                search_row = tk.Frame(self.text_area, bg=BG_COLOR)
                lbl_search = tk.Label(search_row, text=t["lbl_quase_search"], bg=BG_COLOR, fg="white", font=("Consolas", 9))
                lbl_search.pack(side=tk.LEFT, padx=(0, 6))

                self.entry_quase_search = tk.Entry(search_row, textvariable=self.quase_search_var, width=28,
                                                    bg="#333333", fg="white", font=("Consolas", 9), relief=tk.FLAT, insertbackground="white")
                self.entry_quase_search.pack(side=tk.LEFT, ipady=2)
                self.entry_quase_search.bind("<KeyRelease>", lambda event: self.on_quase_search_keyrelease(filepath, filename, event))
                self.entry_quase_search.bind("<Return>", lambda event: self.clear_quase_search(filepath, filename))

                controls_frame = tk.Frame(self.text_area, bg=BG_COLOR)
                btn_cycle = tk.Button(controls_frame, text=self.get_quase_button_label(t), command=lambda: self.cycle_quase_max_count(filepath, filename),
                                      bg="#333333", fg="white", font=("Consolas", 9, "bold"), relief=tk.FLAT, cursor="hand2", padx=8, pady=2)
                btn_cycle.pack(side=tk.LEFT, padx=(0, 6))

                lbl_until = tk.Label(controls_frame, text=t["lbl_quase_or_show_until"], bg=BG_COLOR, fg="white", font=("Consolas", 9))
                lbl_until.pack(side=tk.LEFT, padx=(0, 4))

                vcmd_quase_limit = (self.root.register(lambda p: p == "" or (p.isdigit() and len(p) <= 4)), '%P')
                entry_limit = tk.Entry(controls_frame, textvariable=self.quase_limit_entry_var, validate="key", validatecommand=vcmd_quase_limit,
                                        width=4, bg="#333333", fg="white", font=("Consolas", 9), relief=tk.FLAT, insertbackground="white")
                entry_limit.pack(side=tk.LEFT, padx=(0, 4), ipady=2)
                entry_limit.bind("<Return>", lambda event: self.apply_quase_limit_from_entry(filepath, filename, event))

                btn_fetch = tk.Button(controls_frame, text=t["btn_quase_fetch"], command=lambda: self.apply_quase_limit_from_entry(filepath, filename),
                                      bg="#444444", fg="white", font=("Consolas", 9, "bold"), relief=tk.FLAT, cursor="hand2", padx=8, pady=2)
                btn_fetch.pack(side=tk.LEFT)

                filter_frame = tk.Frame(self.text_area, bg=BG_COLOR)
                for option, label in [("PARTY", t["radio_party"]), ("BOX", t["radio_box"]), ("FAZENDA", t["radio_farm"]), ("TODOS", t["radio_all"])]:
                    tk.Radiobutton(filter_frame, text=label, value=option, variable=self.quase_filter_loc_var,
                                   command=lambda: self.apply_quase_filters(filepath, filename), bg=BG_COLOR, fg="white",
                                   selectcolor=BTN_BG, activebackground=BG_COLOR, activeforeground="white", font=("Consolas", 9),
                                   cursor="hand2").pack(side=tk.LEFT, padx=(0, 10))

                self.chk_quase_wishlist = tk.Checkbutton(
                    filter_frame,
                    text=t["checkbox_wishlist"],
                    variable=self.quase_filter_wishlist_var,
                    command=lambda: self.apply_quase_filters(filepath, filename),
                    bg=BG_COLOR,
                    fg="white",
                    selectcolor=BTN_BG,
                    activebackground=BG_COLOR,
                    activeforeground="white",
                    font=("Consolas", 9),
                    cursor="hand2",
                )
                                
                self.chk_quase_protected = tk.Checkbutton(
                    filter_frame,
                    text=t["checkbox_protected"],
                    variable=self.quase_filter_protected_var,
                    command=lambda: self.apply_quase_filters(filepath, filename),
                    bg=BG_COLOR,
                    fg="white",
                    selectcolor=BTN_BG,
                    activebackground=BG_COLOR,
                    activeforeground="white",
                    font=("Consolas", 9),
                    cursor="hand2",
                )

                self.chk_quase_wishlist.pack(side=tk.LEFT, padx=(0, 10))
                self.chk_quase_wishlist.bind("<Enter>", lambda event: self._show_quase_wishlist_hint(event, self._quase_wishlist_hint_text))
                self.chk_quase_wishlist.bind("<Leave>", self._hide_quase_wishlist_hint)
                self.chk_quase_protected.pack(side=tk.LEFT, padx=(0, 10))

                # ========================================================
                # RADIOBUTTONS (OptionMenu) DE ORDENAÇÃO NA MESMA LINHA
                # ========================================================
                # sort_frame = tk.Frame(self.text_area, bg=BG_COLOR)
                # self.quase_sort_var = tk.StringVar(value=getattr(self, 'quase_sort_by', 'XP'))

                # def apply_sort():
                #     self.quase_sort_by = self.quase_sort_var.get()
                #     self.save_config()
                #     self.process_save(filepath, filename)

                # tk.Radiobutton(sort_frame, text=t.get("radio_sort_xp", "Menor XP Restante"), variable=self.quase_sort_var, value="XP",
                #                command=apply_sort, bg=BG_COLOR, fg="white", selectcolor=BTN_BG, activebackground=BG_COLOR, activeforeground="white", font=("Consolas", 9), cursor="hand2").pack(side=tk.LEFT, padx=(0, 10))
                # tk.Radiobutton(sort_frame, text=t.get("radio_sort_acc", "Menor Talento ACC"), variable=self.quase_sort_var, value="ACC",
                #                command=apply_sort, bg=BG_COLOR, fg="white", selectcolor=BTN_BG, activebackground=BG_COLOR, activeforeground="white", font=("Consolas", 9), cursor="hand2").pack(side=tk.LEFT)

                # ========================================================
                # COMBOBOX (OptionMenu) DE ORDENAÇÃO NA MESMA LINHA
                # ========================================================
                lbl_sort = tk.Label(filter_frame, text="Ordenar:", bg=BG_COLOR, fg="white", font=("Consolas", 9))
                lbl_sort.pack(side=tk.LEFT, padx=(5, 2))

                opt_xp = t.get("radio_sort_xp", "Menor XP Restante")
                opt_acc = t.get("radio_sort_acc", "Menor Talento ACC")
                
                current_sort = opt_acc if getattr(self, 'quase_sort_by', 'XP') == "ACC" else opt_xp
                self.quase_sort_var = tk.StringVar(value=current_sort)

                def on_sort_change(*args):
                    if self.quase_sort_var.get() == opt_acc:
                        self.quase_sort_by = "ACC"
                    else:
                        self.quase_sort_by = "XP"
                    self.save_config()
                    self.process_save(filepath, filename)
                
                self.quase_sort_var.trace("w", on_sort_change)

                sort_menu = tk.OptionMenu(filter_frame, self.quase_sort_var, opt_xp, opt_acc)
                sort_menu.config(bg="#333333", fg="white", font=("Consolas", 9), relief=tk.FLAT, activebackground=BG_COLOR, activeforeground="white", highlightthickness=0, cursor="hand2")
                sort_menu["menu"].config(bg="#333333", fg="white", font=("Consolas", 9))
                sort_menu.pack(side=tk.LEFT)
                # ========================================================
                
                self.text_area.config(state=tk.NORMAL)
                self.text_area.window_create(tk.END, window=search_row)
                self.text_area.insert(tk.END, "\n")
                self.text_area.window_create(tk.END, window=controls_frame)
                self.text_area.insert(tk.END, "\n")
                self.text_area.window_create(tk.END, window=filter_frame)
                self.text_area.insert(tk.END, "\n")
                # self.text_area.window_create(tk.END, window=sort_frame)
                self.text_area.insert(tk.END, "\n")
                self.text_area.config(state=tk.DISABLED)

                combined_entries = []
                for item in almost_list + remaining_list:
                    # 4. ADICIONADO: asc_talent + slot no desempacotamento e no dicionário
                    name, level, talent, faltam_int, bar_str, faltam_str, loc, protected, uid, is_almost, asc_talent, slot = item
                    combined_entries.append({
                        "name": name,
                        "level": level,
                        "goal_level": talent,
                        "remaining": faltam_int,
                        "bar_str": bar_str,
                        "remaining_str": faltam_str,
                        "loc": loc,
                        "slot": slot,
                        "protected": protected,
                        "uid_sort": int(uid, 16) if uid else 0,
                        "is_almost": is_almost,
                        "is_wishlist": False,
                        "wishlist_idx": None,
                        "asc_talent": asc_talent # <-- Adicionado
                    })

                # 5. ADICIONADO: asc_talent + slot no desempacotamento e no dicionário (Wishlist)
                for faltam_int, w_idx, name, level, target_lvl, bar_str, faltam_str, loc, protected, asc_talent, slot in wishlist_pending:
                    combined_entries.append({
                        "name": name,
                        "level": level,
                        "goal_level": target_lvl,
                        "remaining": faltam_int,
                        "bar_str": bar_str,
                        "remaining_str": faltam_str,
                        "loc": loc,
                        "slot": slot,
                        "protected": protected,
                        "uid_sort": w_idx,
                        "is_almost": False,
                        "is_wishlist": True,
                        "wishlist_idx": w_idx,
                        "asc_talent": asc_talent # <-- Adicionado
                    })

                # 6. ADICIONADO: Lógica condicional de ordenação
                if getattr(self, 'quase_sort_by', 'XP') == "ACC":
                    combined_entries.sort(key=lambda x: (x["asc_talent"], x["remaining"], -x["goal_level"], x["name"].lower(), 0 if x["protected"] else 1, x["uid_sort"]))
                else:
                    combined_entries.sort(key=lambda x: (x["remaining"], -x["goal_level"], x["name"].lower(), 0 if x["protected"] else 1, x["uid_sort"]))

                search_text = getattr(self, 'quase_search_text', '').strip().lower()
                filtered = []
                for entry in combined_entries:
                    if self.quase_filter_location != "TODOS" and self.quase_filter_location != entry["loc"]:
                        continue
                    if self.quase_filter_protected and not entry["protected"]:
                        continue
                    if not self.quase_filter_wishlist and entry["is_wishlist"]:
                        continue
                    if search_text and search_text not in entry["name"].lower():
                        continue
                    filtered.append(entry)

                if self.quase_list_max_count > 0:
                    filtered = filtered[:self.quase_list_max_count]

                if filtered:
                    self.text_area.config(state=tk.NORMAL)
                    for entry in filtered:
                        cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[entry["loc"]]
                        lock_icon = "🔒" if entry["protected"] else "  "
                        goal_label = t["target_abbr"] if entry["is_wishlist"] else t["limite_abbr"]
                        self.text_area.insert(tk.END, f" [{self._loc_display(entry['loc'], entry.get('slot')):^7}] ", cor_tag)
                        
                        
                        # --- TRUNCAMENTO INTELIGENTE ---
                        raw_name = entry['name']
                        name_str = raw_name if len(raw_name) <= MAX_NAME_LEN else raw_name[:MAX_NAME_LEN - 3] + "..."
                        # -------------------------------
        
                        # NOVA FORMATAÇÃO: T.A simplificado
                        acc_formatted = f"{{{int(entry['asc_talent']):,} T.A}}".replace(',', '.')
                        
                        # NOVA FORMATAÇÃO: XP com 10 espaços e separador |
                        line_text = f"{lock_icon} {name_str:<{MAX_NAME_LEN}} ({t['lvl_abbr']} {entry['level']:02d} / {entry['goal_level']:02d}) [{entry['bar_str']}] {entry['remaining_str']:>10} EXP  |  {acc_formatted} "
                        
                        line_tag = "almost" if entry["is_almost"] else "status"
                        self.text_area.insert(tk.END, line_text, line_tag)
                        
                        if entry["is_wishlist"] and entry.get("wishlist_idx") is not None:
                            btn_del = tk.Button(self.text_area, text=" ❌ ", command=lambda idx=entry["wishlist_idx"]: self.delete_wishlist_item(idx),
                                                bg="#444444", fg="red", font=("Consolas", 8, "bold"),
                                                relief=tk.FLAT, cursor="hand2", padx=2, pady=0)
                            self.text_area.window_create(tk.END, window=btn_del)
                        self.text_area.insert(tk.END, "\n")
                    self.text_area.config(state=tk.DISABLED)
                    self.text_area.yview(1.0)
                else:
                    self.log(t.get('msg_quase_no_results', t.get('msg_all_normal', '    Todos os Digimons estao evoluindo normalmente.')), 'status')

        def render_lv99_list(is_first, is_last):
            if total_lv99 <= 0:
                return
            render_header_with_controls(t["lbl_lv99_title"], "header_green", 'show_lv99', 'btn_show_lv99', 'lv99', is_first, is_last)

            if getattr(self, 'show_lv99', False):
                for loc in ["PARTY", "BOX", "FAZENDA"]:
                    lv99_list[loc].sort(key=lambda x: (x[0], x[1]))
                    # 8. ADICIONADO: asc_talent + slot no desempacotamento e string
                    for name, level, talent, protected, asc_talent, slot in lv99_list[loc]:
                        cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                        lock_icon = "🔒" if protected else "  "
                        
                        # NOVA FORMATAÇÃO: T.A simplificado
                        acc_formatted = f"{{{int(asc_talent):,} T.A}}".replace(',', '.')
                        
                        self.log([
                            (f" [{self._loc_display(loc, slot):^7}] ", cor_tag),
                            (f"{lock_icon} {name:<{MAX_NAME_LEN}} ({t['lvl_abbr']} {level:02d} / {talent:02d}) {acc_formatted}", "status")
                        ])

        def render_wishlist_list(is_first, is_last):
            if not wishlist_pending:
                return
            wishlist_pending.sort(key=lambda x: x[0])  # Ordena da menor EXP para a maior

            wishlist_title_text = t.get("wishlist_title", " 🎯 WISHLIST / METAS DE EVOLUÇÃO:")
            render_header_with_controls(wishlist_title_text, "header_orange", 'show_wishlist', 'btn_show_wishlist', 'wishlist', is_first, is_last)

            if getattr(self, 'show_wishlist', False):
                # 9. ADICIONADO: asc_talent + slot no desempacotamento e string
                for faltam_int, w_idx, name, level, target_lvl, bar_str, faltam_str, loc, protected, asc_talent, slot in wishlist_pending:
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    lock_icon = "🔒" if protected else "  "

                    self.text_area.config(state=tk.NORMAL)
                    self.text_area.insert(tk.END, f" [{self._loc_display(loc, slot):^7}] ", cor_tag)

                    # NOVA FORMATAÇÃO: T.A simplificado
                    acc_formatted = f"{{{int(asc_talent):,} T.A}}".replace(',', '.')
                    
                    # NOVA FORMATAÇÃO: XP com 10 espaços e separador |
                    raw_name = name
                    name_str = raw_name if len(raw_name) <= MAX_NAME_LEN else raw_name[:MAX_NAME_LEN - 3] + "..."
                    msg = f"{lock_icon} {name_str:<{MAX_NAME_LEN}} ({t['lvl_abbr']} {level:02d} / {target_lvl:02d}) [{bar_str}] {faltam_str:>10} EXP  |  {acc_formatted} "
                    self.text_area.insert(tk.END, msg, "status")

                    btn_del = tk.Button(self.text_area, text=" ❌ ", command=lambda idx=w_idx: self.delete_wishlist_item(idx),
                                        bg="#444444", fg="red", font=("Consolas", 8, "bold"), 
                                        relief=tk.FLAT, cursor="hand2", padx=2, pady=0)
                    
                    self.text_area.window_create(tk.END, window=btn_del)
                    self.text_area.insert(tk.END, "\n")
                    self.text_area.config(state=tk.DISABLED)

        def render_reached_cap_section():
            if not reached_rows:
                return
            self.text_area.config(state=tk.NORMAL)
            self.text_area.insert(tk.END, t.get("cap_list_title", " 🏁 REACHED THE CAP:"), "header_red")
            self.text_area.insert(tk.END, "\n\n")
            self.text_area.config(state=tk.DISABLED)
            
            # 10. ADICIONADO: asc_talent + slot no desempacotamento de entry_data e nas strings
            for entry_type, loc, *entry_data in reached_rows:
                if entry_type == "alert":
                    name, level, talent, protected, asc_talent, slot = entry_data
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    lock_icon = "🔒" if protected else "  "
                    
                   # NOVA FORMATAÇÃO: T.A simplificado
                    acc_formatted = f"{{{int(asc_talent):,} T.A}}".replace(',', '.')
                    
                    self.log([
                        (f" [{self._loc_display(loc, slot):^7}] ", cor_tag),
                        (f"{lock_icon} {name:<{MAX_NAME_LEN}} ({t['lvl_abbr']} {level:02d} / {talent:02d}) {acc_formatted} {txt_atingiu}", "alert")
                    ])
                else:
                    w_idx, w_name, w_level, w_target, w_protected, asc_talent, slot = entry_data
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    w_lock_icon = "🔒" if w_protected else "  "
                    
                    # NOVA FORMATAÇÃO: T.A simplificado
                    w_acc_formatted = f"{{{int(asc_talent):,} T.A}}".replace(',', '.')
                    
                    self.text_area.config(state=tk.NORMAL)
                    self.text_area.insert(tk.END, f" [{self._loc_display(loc, slot):^7}] ", cor_tag)
                    self.text_area.insert(tk.END, f"{w_lock_icon} {w_name:<{MAX_NAME_LEN}} ({t['lvl_abbr']} {w_level:02d} / {w_target:02d}) {w_acc_formatted} {txt_atingiu} ", "status")
                    btn_del = tk.Button(self.text_area, text=" ❌ ", command=lambda idx=w_idx: self.delete_wishlist_item(idx),
                                        bg="#444444", fg="red", font=("Consolas", 8, "bold"),
                                        relief=tk.FLAT, cursor="hand2", padx=2, pady=0)
                    self.text_area.window_create(tk.END, window=btn_del)
                    self.text_area.insert(tk.END, "\n")
                    self.text_area.config(state=tk.DISABLED)

        renderers = {
            'almost': render_almost_list,
            'lv99': render_lv99_list,
            'wishlist': render_wishlist_list,
        }

        order = self.get_normalized_list_order()
        sections_rendered = 0
        if reached_rows:
            render_reached_cap_section()
            sections_rendered += 1

        for idx, key in enumerate(order):
            if key == 'almost' and (len(almost_list) + len(remaining_list)) > 0:
                if sections_rendered > 0:
                    append_section_divider()
                renderers[key](idx == 0, idx == len(order) - 1)
                sections_rendered += 1
            elif key == 'lv99' and total_lv99 > 0:
                if sections_rendered > 0:
                    append_section_divider()
                renderers[key](idx == 0, idx == len(order) - 1)
                sections_rendered += 1
            elif key == 'wishlist' and wishlist_pending:
                if sections_rendered > 0:
                    append_section_divider()
                renderers[key](idx == 0, idx == len(order) - 1)
                sections_rendered += 1

        if not self.is_paused:
            self.log(t["waiting_msg"], "status")
        else:
            self.log(t["paused_msg"], "alert")

    def update_loop(self):
        if self.is_paused:
            self.blink_state = not self.blink_state
            new_bg = "#FF0000" if self.blink_state else "#8B0000"
            self.btn_resume.config(bg=new_bg)
        else:
            try:
                latest_file, current_mtime = self.get_target_save()
                if latest_file and current_mtime != self.last_mtime:
                    # DEBOUNCE / "ESTABILIZAÇÃO" DO ARQUIVO
                    #
                    # Antes, o mtime mudar já disparava process_save() (e, com o
                    # auto-sync ligado, também o sync com o Excel) NO MESMO
                    # instante. Se o jogo ainda estivesse no meio da escrita do
                    # save (arquivo aberto, conteúdo parcial), a gente podia:
                    #   1) ler um save truncado/corrompido (openssl decodifica
                    #      lixo, sem erro nenhum - garbage in, garbage out), e/ou
                    #   2) competir por acesso ao MESMO arquivo (nosso openssl.exe
                    #      + o processo do jogo escrevendo), o que em alguns jogos
                    #      pode causar erro de I/O na hora de salvar.
                    #
                    # Isso é ainda mais arriscado com o auto-sync ligado, porque
                    # ele soma trabalho extra (chamadas COM pro Excel) bem na
                    # janela em que o jogo pode estar salvando.
                    #
                    # A correção: só processamos o arquivo depois de ver o MESMO
                    # (mtime, tamanho) em duas checagens seguidas (~2s de
                    # intervalo, o próprio ciclo do update_loop) - ou seja, só
                    # depois que ele parou de mudar. Isso não deixa a detecção
                    # mais lenta em uso normal (o jogo grava rápido e para), só
                    # evita agir enquanto o arquivo ainda está sendo escrito.
                    try:
                        current_size = os.path.getsize(latest_file)
                    except OSError:
                        current_size = None

                    candidate = (latest_file, current_mtime, current_size)

                    if self._pending_save_check == candidate:
                        # Mesmo arquivo, mesmo mtime, mesmo tamanho do ciclo
                        # anterior -> parou de mudar, seguro processar agora.
                        self._pending_save_check = None
                        self.last_mtime = current_mtime
                        filename_only = os.path.basename(latest_file)
                        self.process_save(latest_file, filename_only)
                        self.maybe_auto_sync_talentos(filename_only, current_mtime)
                    else:
                        # 1ª vez vendo essa mudança (ou o arquivo ainda está
                        # mudando) - só guarda e confirma no próximo ciclo.
                        self._pending_save_check = candidate
            except Exception:
                pass 
        self.root.after(2000, self.update_loop)

if __name__ == "__main__":
    root = tk.Tk()
    app = DigimonMonitorApp(root)
    if app.save_dir:
        root.mainloop()