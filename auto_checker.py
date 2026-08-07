import os
import time
import struct
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json

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

# CORES DARK MODE
BG_COLOR = "#121212"
FG_COLOR = "#00FF00"  
FG_ALERT = "#FF4500"  
FG_ALMOST = "#FFD700" 
PANEL_BG = "#1E1E1E"
BTN_BG = "#333333"

# ==========================================
# DICIONÁRIO DE IDIOMAS (i18n)
# ==========================================
I18N = {
    "EN": {
        "settings": "⚙️ SETTINGS",
        "atingiu": "<-- REACHED THE CAP!",
        "cap_list_title": " 🏁 REACHED THE CAP:",
        "quase": " ⏳ ALMOST THERE (Missing 1 Level):",
        "modo_auto": "Automatic Mode",
        "modo_manual": "Manual Mode (Ignore auto-save)",
        "btn_change_folder": "📂 Change Save Folder",
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
        "protected_list_title": " 🔒 PROTECTED DIGIMONS:",
        "btn_show_protected": "Click to see protected digimons",
        "btn_hide_details": "Hide details",
        "btn_show_lv99": "Click to see the Level 99 digimons",
        "btn_show_almost": "Click to see who's almost there",
        "btn_show_wishlist": "Click to see the Wishlist",
        "lbl_lv99_title": " 👑 DIGIMONS AT MAXIMUM LEVEL (99):",
        "waiting_msg": "\n[Waiting for game update... Keep this open on the 2nd screen]",
        "paused_msg": "\n[TRACKING PAUSED. Click the red side button to return to the radar]",
        "lvl_abbr": "Lv.",
        "limite_abbr": "Cap",
        "faltam_abbr": "Missing:",
        "wishlist_title": " 🎯 TARGET WISHLIST / EVOLUTION GOALS:",
        "btn_view_remaining": "View remaining Digimons",
        "btn_hide_remaining": "Hide remaining Digimons",
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
        "wishlist_not_found_msg": "'{name}' was not found in the current save. It may have evolved, been released, or you might be viewing a different save file.",
        "wishlist_readded_msg": "'{name}' is present in the current save again and was restored to the wishlist.",
        "target_abbr": "Target",
        "loc_labels": {"PARTY": "PARTY", "BOX": "BOX", "FAZENDA": "FARM"}
    },
    "PT": {
        "settings": "⚙️ CONFIGURAÇÕES",
        "atingiu": "<-- ATINGIU O LIMITE!",
        "cap_list_title": " 🏁 CHEGOU NO CAP:",
        "quase": " ⏳ QUASE LÁ (Faltando 1 Level):",
        "modo_auto": "Modo Automático",
        "modo_manual": "Modo Manual (Ignora o Auto-Save)",
        "btn_change_folder": "📂 Mudar Pasta de Saves",
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
        "protected_list_title": " 🔒 DIGIMONS PROTEGIDOS:",
        "btn_show_protected": "Clique para ver os digimons protegidos",
        "btn_hide_details": "Ocultar detalhes",
        "btn_show_lv99": "Clique para ver os digimons no Nv. 99",
        "btn_show_almost": "Clique para ver quem está quase lá",
        "btn_show_wishlist": "Clique para ver a Wishlist",
        "lbl_lv99_title": " 👑 DIGIMONS NO LEVEL MÁXIMO (99):",
        "waiting_msg": "\n[Aguardando atualização do jogo... Mantenha aberto na 2ª tela]",
        "paused_msg": "\n[RASTREAMENTO PAUSADO. Clique no botão lateral vermelho para voltar ao radar]",
        "lvl_abbr": "Nv.",
        "limite_abbr": "Limite",
        "faltam_abbr": "Faltam:",
        "wishlist_title": " 🎯 WISHLIST / METAS DE EVOLUÇÃO:",
        "btn_view_remaining": "Visualizar o restante dos digimons",
        "btn_hide_remaining": "Ocultar Digimons restantes",
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
        "wishlist_not_found_msg": "'{name}' não foi encontrado no save atual. Ele pode ter evoluído, sido liberado, ou você pode estar vendo um save diferente.",
        "wishlist_readded_msg": "'{name}' está presente no save atual novamente e foi restaurado na wishlist.",
        "target_abbr": "Alvo",
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
        # Estado temporário da sublista 'restantes' (não persistido no config.json)
        self.show_remaining = False
            
        self.mode = tk.StringVar(value="AUTO")
        self.is_paused = False
        self.last_mtime = 0
        self.blink_state = False

        self.setup_ui()
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
        self.show_protected = self.config_data.get("show_protected", False)

        # Ordem das listas reordenáveis (2, 3, 4, 5). A 1ª lista (Level Cap) é sempre fixa.
        self.list_order = self.config_data.get("list_order", ["almost", "lv99", "wishlist", "protected"])

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
        self.config_data["show_protected"] = getattr(self, 'show_protected', False)
        self.config_data["list_order"] = self.get_normalized_list_order()
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config_data, f, indent=4)
        except Exception:
            pass

    def on_closing(self):
        self.save_config()
        self.root.destroy()

    def get_normalized_list_order(self):
        """Garante que self.list_order sempre contenha exatamente as 4 chaves válidas
        (almost, lv99, wishlist, protected), preservando a ordem já salva e anexando
        no final qualquer chave nova que ainda não existisse (ex: quando esta lista
        de Protegidos foi adicionada em uma versão mais nova do programa)."""
        valid_keys = ["almost", "lv99", "wishlist", "protected"]
        current = getattr(self, 'list_order', None) or []
        normalized = [k for k in current if k in valid_keys]
        for k in valid_keys:
            if k not in normalized:
                normalized.append(k)
        self.list_order = normalized
        return normalized

    def move_list_order(self, key, direction):
        """Move uma lista reordenável (2ª a 5ª) uma posição pra cima (-1) ou pra baixo (+1)."""
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

        if hasattr(self, "_current_filepath") and os.path.exists(self._current_filepath):
            self.process_save(self._current_filepath, self._current_filename)

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
        
        tk.Radiobutton(lang_frame, text="🇧🇷 Português", variable=self.lang_var, value="PT", 
                       command=self.on_lang_change, bg=PANEL_BG, fg="white", selectcolor=BTN_BG, 
                       font=("Consolas", 10)).pack(side=tk.LEFT, expand=True)
                       
        tk.Radiobutton(lang_frame, text="🇺🇸 English", variable=self.lang_var, value="EN", 
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

        # 1. Busca por ID ou Nome (label em cima, campo embaixo, botão de largura total)
        self.lbl_wishlist_search = tk.Label(control_frame, text=I18N[self.lang]["wishlist_search_label"], bg=PANEL_BG, fg="white", font=("Consolas", 9))
        self.lbl_wishlist_search.pack(anchor="w", padx=15, pady=(0, 2))

        self.entry_wish_id = tk.Entry(control_frame, bg="#333333", fg="white", font=("Consolas", 9), relief=tk.FLAT, insertbackground="white")
        self.entry_wish_id.pack(fill='x', padx=15, pady=(0, 4), ipady=2)

        # ATALHO: Funcionar ao pressionar Enter no teclado
        self.entry_wish_id.bind("<Return>", lambda event: self.search_wishlist_digimon())

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
        self.text_area.tag_config("status", foreground="white")
        self.text_area.tag_config("loc_party", foreground="#FFA500") 
        self.text_area.tag_config("loc_box", foreground="#1E90FF")   
        self.text_area.tag_config("loc_fazenda", foreground="#32CD32") 

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
        1 = protegido, 0 = desprotegido."""
        offset = name_offset + 0x11C
        if offset + 4 > len(data):
            return False
        return struct.unpack_from("<I", data, offset)[0] == 1

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

    def search_wishlist_digimon(self):
        """Busca o Digimon pelo ID ou pelo Nome no save atualmente ativo."""
        t = I18N[self.lang]
        loc_lbl = t["loc_labels"]
        query_text = self.entry_wish_id.get().strip()
        if not query_text:
            messagebox.showwarning(t["msg_warning_title"], t["wishlist_err_empty_query"])
            return

        is_id_search = query_text.isdigit()
        search_id = int(query_text) if is_id_search else None
        search_name = query_text.lower() if not is_id_search else None

        if not os.path.exists(SAVE_FILE_DEC):
            messagebox.showwarning(t["msg_warning_title"], t["wishlist_err_no_save"])
            return

        with open(SAVE_FILE_DEC, "rb") as f:
            data = f.read()

        self.search_results_map = []
        combo_options = []

        def match_criteria(d_id, d_name):
            if is_id_search:
                return d_id == search_id
            return search_name in d_name.lower()

        # 1. Busca na Fazenda
        FARM_START = 0x539C8
        FARM_SIZE = 344
        HEADER_FARM = 0x18
        for i in range(30):
            offset = FARM_START + (i * FARM_SIZE)
            if offset + FARM_SIZE <= len(data) and data[offset] == 1:
                name_offset = offset + HEADER_FARM
                d_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
                
                name_bytes = bytearray()
                for b in data[name_offset : name_offset + 32]:
                    if b == 0: break
                    name_bytes.append(b)
                name = name_bytes.decode('ascii', errors='ignore')

                if match_criteria(d_id, name):
                    uid = self.read_digimon_uid(data, name_offset)
                    level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                    exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                    talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]
                    talent = talent_raw // 1000 if talent_raw >= 1000 else 1
                    if talent > 99: talent = 99
                    protected = self.read_digimon_protected(data, name_offset)
                    
                    info = {'uid': uid, 'id': d_id, 'name': name, 'loc': 'FAZENDA', 'slot': i, 'level': level, 'exp': exp, 'talent': talent, 'protected': protected}
                    self.search_results_map.append(info)
                    exp_str = self.get_level_exp_display(d_id, level, exp)
                    lock_icon = " 🔒" if protected else ""
                    combo_options.append(f"{name}{lock_icon} [{loc_lbl['FAZENDA']}] - {t['lvl_abbr']}{level}/{talent} | EXP {exp_str} | Ref {uid}")

        # 2. Busca na Party e Box
        regions = [("PARTY", 0x12C8 - 0x10, 0x10, 6), ("BOX", 0x1AA8 - 0x10, 0x10, 999)]
        for loc, start, h_size, max_s in regions:
            for i in range(max_s):
                offset = start + (i * DIGIMON_SIZE)
                if offset + DIGIMON_SIZE <= len(data) and data[offset] == 1:
                    name_offset = offset + h_size
                    d_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
                    
                    name_bytes = bytearray()
                    for b in data[name_offset : name_offset + 32]:
                        if b == 0: break
                        name_bytes.append(b)
                    name = name_bytes.decode('ascii', errors='ignore')

                    if match_criteria(d_id, name):
                        uid = self.read_digimon_uid(data, name_offset)
                        level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                        exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                        talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]
                        talent = talent_raw // 1000 if talent_raw >= 1000 else 1
                        if talent > 99: talent = 99
                        protected = self.read_digimon_protected(data, name_offset)
                        
                        info = {'uid': uid, 'id': d_id, 'name': name, 'loc': loc, 'slot': i, 'level': level, 'exp': exp, 'talent': talent, 'protected': protected}
                        self.search_results_map.append(info)
                        exp_str = self.get_level_exp_display(d_id, level, exp)
                        lock_icon = " 🔒" if protected else ""
                        combo_options.append(f"{name}{lock_icon} [{loc_lbl[loc]}] - {t['lvl_abbr']}{level}/{talent} | EXP {exp_str} | Ref {uid}")

        if self.search_results_map:
            def sort_wishlist_result(info):
                level_exp, _ = self.get_level_exp_progress(info['id'], info['level'], info['exp'])
                return (info['level'], level_exp, info['talent'])

            self.search_results_map.sort(key=sort_wishlist_result, reverse=True)
            combo_options = []
            for info in self.search_results_map:
                exp_str = self.get_level_exp_display(info['id'], info['level'], info['exp'])
                lock_icon = " 🔒" if info.get('protected') else ""
                combo_options.append(f"{info['name']}{lock_icon} [{loc_lbl[info['loc']]}] - {t['lvl_abbr']}{info['level']}/{info['talent']} | EXP {exp_str} | Ref {info['uid']}")

            self.combo_wish_results['values'] = combo_options
            self.combo_wish_results.current(0)
            result_word = "resultado" if len(combo_options) == 1 else "resultados"
            self.lbl_wishlist_result_count.config(text=f"{len(combo_options)} {result_word}")
        else:
            self.combo_wish_results['values'] = []
            self.combo_wish_results.set("")
            self.lbl_wishlist_result_count.config(text="0 resultados")
            messagebox.showinfo(t["msg_search_title"], t["wishlist_no_results"].format(query=query_text))

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
            'orphaned': False
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

    def update_summary_panel(self, count_party, count_box, count_farm, total_alerts, total_almost, total_lv99, total_protected, t, loc_lbl):
        """Preenche o painel de Resumo na barra lateral: ocupação de Party/Box/Fazenda (labels grandes) + resumo de pendências (texto)."""
        self.lbl_party_count.config(text=f"{loc_lbl['PARTY']}: {count_party}/6")
        self.lbl_box_count.config(text=f"{loc_lbl['BOX']}: {count_box}/999")
        self.lbl_farm_count.config(text=f"{loc_lbl['FAZENDA']}: {count_farm}/30")

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

    def on_mode_change(self):
        self.is_paused = False
        self.lbl_paused.pack_forget()
        self.btn_resume.pack_forget()
        self.last_mtime = 0 
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
        self.btn_resume.config(bg="#8B0000") 
        self.save_combo.selection_clear()
        self.root.focus_set()

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
        
    def calculate_almost_data(self, data, name_offset, name, level, talent, loc, protected):
        digimon_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
        current_exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
        
        exp_alvo = get_exp_needed(digimon_id, talent)
        exp_base = get_exp_needed(digimon_id, level)
        
        faltam = exp_alvo - current_exp
        
        progresso_total = exp_alvo - exp_base
        progresso_atual = current_exp - exp_base
        
        porcentagem = 0
        if progresso_total > 0:
            porcentagem = max(0, min(1, progresso_atual / progresso_total))
            
        bar_length = 10
        filled = int(porcentagem * bar_length)
        bar_str = ("█" * filled) + ("░" * (bar_length - filled))
        
        faltam_str = f"{faltam:,}".replace(",", ".")
        
        return (name, level, talent, faltam, bar_str, faltam_str, loc, protected)

    def process_save(self, filepath, filename):
        self._current_filepath = filepath
        self._current_filename = filename
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
        protected_list = []
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

                digimon_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
                uid = self.read_digimon_uid(data, name_offset)
                level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                current_exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]
                protected = self.read_digimon_protected(data, name_offset)

                talent = talent_raw // 1000 if talent_raw >= 1000 else 1
                if talent > 99: talent = 99

                active_digimons[("FAZENDA", i)] = {
                    'uid': uid, 'id': digimon_id, 'name': name, 'level': level, 'exp': current_exp, 'loc': "FAZENDA", 'slot': i, 'protected': protected, 'talent': talent
                }

                if talent_raw >= 1000:
                    talent = talent_raw // 1000
                    if talent > 99: talent = 99 
                    
                    if level == 99:
                        lv99_list["FAZENDA"].append((name, level, talent, protected))
                        total_lv99 += 1
                    elif level >= talent: 
                        alerts["FAZENDA"].append((name, level, talent, protected))
                        total_alerts += 1
                    elif level == talent - 1:
                        almost_data = self.calculate_almost_data(data, name_offset, name, level, talent, "FAZENDA", protected)
                        almost_list.append(almost_data)
                        total_almost += 1

                    if protected:
                        p_name, p_level, p_talent, p_faltam, p_bar, p_faltam_str, p_loc, p_prot = self.calculate_almost_data(data, name_offset, name, level, talent, "FAZENDA", protected)
                        if p_faltam <= 0:
                            p_faltam, p_faltam_str, p_bar = 0, "0", "█" * 10
                        protected_list.append((p_name, p_level, p_talent, p_faltam, p_bar, p_faltam_str, p_loc, p_prot))
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

                digimon_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
                uid = self.read_digimon_uid(data, name_offset)
                level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                current_exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]
                protected = self.read_digimon_protected(data, name_offset)

                talent = talent_raw // 1000 if talent_raw >= 1000 else 1
                if talent > 99: talent = 99

                active_digimons[(loc, i)] = {
                    'uid': uid, 'id': digimon_id, 'name': name, 'level': level, 'exp': current_exp, 'loc': loc, 'slot': i, 'protected': protected, 'talent': talent
                }

                if talent_raw >= 1000:
                    talent = talent_raw // 1000
                    if talent > 99: talent = 99 
                    
                    if level == 99:
                        lv99_list[loc].append((name, level, talent, protected))
                        total_lv99 += 1
                    elif level >= talent: 
                        alerts[loc].append((name, level, talent, protected))
                        total_alerts += 1
                    elif level == talent - 1:
                        almost_data = self.calculate_almost_data(data, name_offset, name, level, talent, loc, protected)
                        almost_list.append(almost_data)
                        total_almost += 1

                    if protected:
                        p_name, p_level, p_talent, p_faltam, p_bar, p_faltam_str, p_loc, p_prot = self.calculate_almost_data(data, name_offset, name, level, talent, loc, protected)
                        if p_faltam <= 0:
                            p_faltam, p_faltam_str, p_bar = 0, "0", "█" * 10
                        protected_list.append((p_name, p_level, p_talent, p_faltam, p_bar, p_faltam_str, p_loc, p_prot))
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
                if item.get('uid') != dig_info.get('uid'):
                    item['uid'] = dig_info.get('uid', '')
                    wishlist_state_changed = True
                if item.get('loc') != dig_info['loc'] or item.get('slot') != dig_info['slot']:
                    item['loc'] = dig_info['loc']
                    item['slot'] = dig_info['slot']
                    wishlist_state_changed = True
                wishlist_resolved.append((w_idx, item, dig_info))
            else:
                # Digimon sumiu do save (evoluiu, foi deletado, ou o save mudou) -> vira órfão
                item['orphaned'] = True
                wishlist_orphaned.append((w_idx, item))
                wishlist_state_changed = True

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

            exp_target = get_exp_needed(dig_id, target_lvl)
            exp_base = get_exp_needed(dig_id, cur_lvl)
            faltam = exp_target - cur_exp

            if faltam <= 0 or cur_lvl >= target_lvl:
                wishlist_reached.append((w_idx, w_name, cur_lvl, target_lvl, w_loc, w_protected))
            else:
                prog_total = exp_target - exp_base
                prog_atual = cur_exp - exp_base
                pct = prog_atual / prog_total if prog_total > 0 else 0
                pct = max(0.0, min(1.0, pct))
                filled = int(pct * 10)
                bar_str = ("█" * filled) + ("░" * (10 - filled))
                faltam_str = f"{faltam:,}".replace(",", ".")
                wishlist_pending.append((faltam, w_idx, w_name, cur_lvl, target_lvl, bar_str, faltam_str, w_loc, w_protected))

        # Os que bateram a meta contam no resumo junto com os alertas normais
        total_alerts += len(wishlist_reached)

        self.save_combo.set(filename)
        self.save_combo.selection_clear()
        
        self.log("=" * 75, "header")
        self.log(t["app_title"], "header")
        self.log("=" * 75, "header")
        
        status_msg = f"{t['status_inspecting']}{filename}" if self.is_paused else f"{t['status_monitoring']}{filename}"
        self.log(f"{t['status_prefix']}{status_msg}{t['lbl_updated']}{time.strftime('%H:%M:%S')}", "status")
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
            for name, level, talent, protected in alerts[loc]:
                reached_rows.append(("alert", loc, name, level, talent, protected))

            for w_idx, w_name, w_level, w_target, w_item_loc, w_protected in wishlist_reached:
                if w_item_loc != loc:
                    continue
                reached_rows.append(("wishlist", loc, w_idx, w_name, w_level, w_target, w_protected))

        self.update_summary_panel(count_party, count_box, count_farm, total_alerts, total_almost, total_lv99, total_protected, t, loc_lbl)

        # ==========================================
        # LISTAS 2 a 5 — ORDEM CUSTOMIZÁVEL PELO USUÁRIO (setas ▲▼ no cabeçalho)
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
            if total_almost <= 0:
                return
            render_header_with_controls(txt_quase, "almost", 'show_almost', 'btn_show_almost', 'almost', is_first, is_last)

            if getattr(self, 'show_almost', False):
                almost_list.sort(key=lambda x: x[3])
                for name, level, talent, faltam_int, bar_str, faltam_str, loc, protected in almost_list:
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    lock_icon = "🔒" if protected else "  "
                    self.log([
                        (f" [{loc_lbl[loc]:^7}] ", cor_tag),
                        (f"{name:<12}{lock_icon} ({t['lvl_abbr']} {level:02d} / {t['limite_abbr']} {talent:02d}) [{bar_str}] {t['faltam_abbr']} {faltam_str:>7} EXP", "status")
                    ])

                # Botão para visualizar/ocultar a sublista enorme dos restantes
                def toggle_remaining():
                    self.show_remaining = not getattr(self, 'show_remaining', False)
                    # Re-renderiza a tela atual
                    self.process_save(filepath, filename)

                btn_label = t.get('btn_view_remaining') if not getattr(self, 'show_remaining', False) else t.get('btn_hide_remaining')
                btn_toggle = tk.Button(self.text_area, text=btn_label, command=toggle_remaining,
                                       bg="#333333", fg="white", font=("Consolas", 8, "bold"),
                                       relief=tk.FLAT, cursor="hand2", padx=8, pady=0)
                # Pula uma linha antes do botão
                self.text_area.insert(tk.END, "\n")
                self.text_area.window_create(tk.END, window=btn_toggle)
                # Pula uma linha depois do botão
                self.text_area.insert(tk.END, "\n\n")

                # Renderiza a sublista com todos os digimons restantes (quando ativa)
                def render_remaining_sublist():
                    if not getattr(self, 'show_remaining', False):
                        return

                    # Construir lista de candidatos: todos os active_digimons que não estão em 'almost' e que ainda estão abaixo do talento
                    remaining = []
                    for key, cand in active_digimons.items():
                        try:
                            d_id = cand['id']
                            lvl = cand['level']
                            cur_exp = cand['exp']
                            talent = cand.get('talent', 1)
                        except Exception:
                            continue

                        # Excluir quem está na lista 'Quase lá' (level == talent - 1) e quem já atingiu/ultrapassou o talento
                        if talent <= 1:
                            continue
                        if lvl == talent - 1:
                            continue
                        if lvl >= talent:
                            continue

                        exp_alvo = get_exp_needed(d_id, talent)
                        exp_base = get_exp_needed(d_id, lvl)
                        faltam = exp_alvo - cur_exp
                        progresso_total = exp_alvo - exp_base
                        progresso_atual = cur_exp - exp_base
                        porcentagem = 0
                        if progresso_total > 0:
                            porcentagem = max(0, min(1, progresso_atual / progresso_total))
                        filled = int(porcentagem * 10)
                        bar_str = ("█" * filled) + ("░" * (10 - filled))
                        faltam_str = f"{faltam:,}".replace(",", ".")
                        remaining.append((faltam, cand['name'], lvl, talent, bar_str, faltam_str, cand['loc'], cand.get('protected', False)))

                    remaining.sort(key=lambda x: (x[0], x[1]))

                    # Exibir sumário e linhas (pode ser grande)
                    self.text_area.config(state=tk.NORMAL)
                    self.text_area.insert(tk.END, f"--- {len(remaining)} digimon(s) restantes ---\n", "status")
                    for faltam, name, lvl, talent, bar_str, faltam_str, loc, protected in remaining:
                        cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                        lock_icon = "🔒" if protected else "  "
                        self.text_area.insert(tk.END, f" [{loc_lbl[loc]:^7}] ", cor_tag)
                        self.text_area.insert(tk.END, f"{name:<12}{lock_icon} ({t['lvl_abbr']} {lvl:02d} / {t['limite_abbr']} {talent:02d}) [{bar_str}] {t['faltam_abbr']} {faltam_str:>7} EXP\n", "status")
                    self.text_area.insert(tk.END, "\n")
                    self.text_area.config(state=tk.DISABLED)

                render_remaining_sublist()

        def render_lv99_list(is_first, is_last):
            if total_lv99 <= 0:
                return
            render_header_with_controls(t["lbl_lv99_title"], "loc_fazenda", 'show_lv99', 'btn_show_lv99', 'lv99', is_first, is_last)

            if getattr(self, 'show_lv99', False):
                for loc in ["PARTY", "BOX", "FAZENDA"]:
                    lv99_list[loc].sort(key=lambda x: (x[0], x[1]))
                    for name, level, talent, protected in lv99_list[loc]:
                        cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                        lock_icon = "🔒" if protected else "  "
                        self.log([
                            (f" [{loc_lbl[loc]:^7}] ", cor_tag),
                            (f"{name:<16}{lock_icon} ({t['lvl_abbr']} {level:02d} / {t['limite_abbr']} {talent:02d})", "status")
                        ])

        def render_wishlist_list(is_first, is_last):
            if not wishlist_pending:
                return
            wishlist_pending.sort(key=lambda x: x[0])  # Ordena da menor EXP para a maior

            wishlist_title_text = t.get("wishlist_title", " 🎯 WISHLIST / METAS DE EVOLUÇÃO:")
            render_header_with_controls(wishlist_title_text, "header", 'show_wishlist', 'btn_show_wishlist', 'wishlist', is_first, is_last)

            if getattr(self, 'show_wishlist', False):
                for faltam_int, w_idx, name, level, target_lvl, bar_str, faltam_str, loc, protected in wishlist_pending:
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    lock_icon = "🔒" if protected else "  "

                    self.text_area.config(state=tk.NORMAL)
                    self.text_area.insert(tk.END, f" [{loc_lbl[loc]:^7}] ", cor_tag)

                    msg = f"{name:<12}{lock_icon} ({t['lvl_abbr']} {level:02d} / {t['target_abbr']} {target_lvl:02d}) [{bar_str}] {t['faltam_abbr']} {faltam_str:>7} EXP "
                    self.text_area.insert(tk.END, msg, "status")

                    # Botão Deletar Inline [X]
                    btn_del = tk.Button(self.text_area, text=" ❌ ", command=lambda idx=w_idx: self.delete_wishlist_item(idx),
                                        bg="#444444", fg="red", font=("Consolas", 8, "bold"), 
                                        relief=tk.FLAT, cursor="hand2", padx=2, pady=0)
                    
                    self.text_area.window_create(tk.END, window=btn_del)
                    self.text_area.insert(tk.END, "\n")
                    self.text_area.config(state=tk.DISABLED)

        def render_protected_list(is_first, is_last):
            if total_protected <= 0:
                return
            protected_list.sort(key=lambda x: x[3])  # Ordena da menor EXP faltando para a maior
            render_header_with_controls(t["protected_list_title"], "header_orange", 'show_protected', 'btn_show_protected', 'protected', is_first, is_last)

            if getattr(self, 'show_protected', False):
                for name, level, talent, faltam_int, bar_str, faltam_str, loc, protected in protected_list:
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    if level >= talent:
                        self.log([
                            (f" [{loc_lbl[loc]:^7}] ", cor_tag),
                            (f"🔒 {name:<12} ({t['lvl_abbr']} {level:02d} / {t['limite_abbr']} {talent:02d}) {txt_atingiu}", "alert")
                        ])
                    else:
                        self.log([
                            (f" [{loc_lbl[loc]:^7}] ", cor_tag),
                            (f"🔒 {name:<12} ({t['lvl_abbr']} {level:02d} / {t['limite_abbr']} {talent:02d}) [{bar_str}] {t['faltam_abbr']} {faltam_str:>7} EXP", "status")
                        ])

        def render_reached_cap_section():
            if not reached_rows:
                return
            self.text_area.config(state=tk.NORMAL)
            self.text_area.insert(tk.END, t.get("cap_list_title", " 🏁 REACHED THE CAP:"), "header_red")
            self.text_area.insert(tk.END, "\n\n")
            self.text_area.config(state=tk.DISABLED)
            for entry_type, loc, *entry_data in reached_rows:
                if entry_type == "alert":
                    name, level, talent, protected = entry_data
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    lock_icon = "🔒" if protected else "  "
                    self.log([
                        (f" [{loc_lbl[loc]:^7}] ", cor_tag),
                        (f"{name:<16}{lock_icon} ({t['lvl_abbr']} {level:02d} / {t['limite_abbr']} {talent:02d}) {txt_atingiu}", "alert")
                    ])
                else:
                    w_idx, w_name, w_level, w_target, w_protected = entry_data
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    w_lock_icon = "🔒" if w_protected else "  "
                    self.text_area.config(state=tk.NORMAL)
                    self.text_area.insert(tk.END, f" [{loc_lbl[loc]:^7}] ", cor_tag)
                    self.text_area.insert(tk.END, f"{w_name:<16}{w_lock_icon} ({t['lvl_abbr']} {w_level:02d} / {t['target_abbr']} {w_target:02d}) {txt_atingiu} ", "status")
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
            'protected': render_protected_list,
        }

        order = self.get_normalized_list_order()
        sections_rendered = 0
        if reached_rows:
            render_reached_cap_section()
            sections_rendered += 1

        for idx, key in enumerate(order):
            if key == 'almost' and total_almost > 0:
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
            elif key == 'protected' and total_protected > 0:
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
                    self.last_mtime = current_mtime
                    filename_only = os.path.basename(latest_file)
                    self.process_save(latest_file, filename_only)
            except Exception:
                pass 
        self.root.after(2000, self.update_loop)

if __name__ == "__main__":
    root = tk.Tk()
    app = DigimonMonitorApp(root)
    if app.save_dir:
        root.mainloop()