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
        "summary_almost": " digimon(s) need 1 more level to reach the Maximum limit",
        "summary_lv99": " digimon(s) are at level 99",
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
        "target_reached": "<-- TARGET REACHED!",
        "wishlist_auto_removed": "⚠️  WISHLIST: digimon(s) not found in the current save (evolved / released / different save):",
        "wishlist_panel_title": "🎯 WISHLIST / TARGET TRACKER",
        "wishlist_search_label": "Search:",
        "wishlist_search_btn": "🔍 Search",
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
        "target_reached": "<-- META ATINGIDA!",
        "wishlist_auto_removed": "⚠️  WISHLIST: digimon(s) não encontrado(s) no save atual (evoluiu / foi liberado / save diferente):",
        "wishlist_panel_title": "🎯 WISHLIST / META TRACKER",
        "wishlist_search_label": "Busca:",
        "wishlist_search_btn": "🔍 Buscar",
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
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config_data, f, indent=4)
        except Exception:
            pass

    def on_closing(self):
        self.save_config()
        self.root.destroy()

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
        self.btn_wish_search.config(text=t["wishlist_search_btn"])
        self.lbl_wishlist_target.config(text=t["wishlist_target_label"])
        self.btn_wish_add.config(text=t["wishlist_add_btn"])

        if hasattr(self, "_current_filepath") and os.path.exists(self._current_filepath):
            self.process_save(self._current_filepath, self._current_filename)

    def setup_ui(self):
        main_frame = tk.Frame(self.root, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True)

        control_frame = tk.Frame(main_frame, bg=PANEL_BG, width=350)
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
        self.btn_wish_search.pack(fill='x', padx=15, pady=(0, 6), ipady=2)

        # 2. Combobox de Resultados Encontrados
        self.combo_wish_results = ttk.Combobox(control_frame, state="readonly", font=("Consolas", 9))
        self.combo_wish_results.pack(fill='x', padx=15, pady=(0, 6))

        # 3. Target Level + Botão Adicionar (lado a lado, já que o campo é só um número de 2 dígitos)
        target_frame = tk.Frame(control_frame, bg=PANEL_BG)
        target_frame.pack(fill='x', padx=15, pady=(0, 10))

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

        self.pause_frame = tk.Frame(control_frame, bg=PANEL_BG)
        self.pause_frame.pack(pady=15, fill=tk.X, padx=15)
        
        self.lbl_paused = tk.Label(self.pause_frame, text=I18N[self.lang]["lbl_paused"], bg=PANEL_BG, fg="yellow", font=("Consolas", 12, "bold"))
        self.btn_resume = tk.Button(self.pause_frame, text=I18N[self.lang]["btn_resume"], command=self.resume_tracking, 
                                    bg="#8B0000", fg="white", font=("Consolas", 10, "bold"), relief=tk.FLAT)

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
        self.text_area.tag_config("status", foreground="white")
        self.text_area.tag_config("loc_party", foreground="#FFA500") 
        self.text_area.tag_config("loc_box", foreground="#1E90FF")   
        self.text_area.tag_config("loc_fazenda", foreground="#32CD32") 

    def _validate_target_lvl_input(self, proposed_value):
        """Validação de campo (usada pelo Entry do Target Level): só aceita vazio ou até 2 dígitos numéricos (nível vai de 1 a 99)."""
        if proposed_value == "":
            return True
        return proposed_value.isdigit() and len(proposed_value) <= 2

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
                    level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                    exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                    talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]
                    talent = talent_raw // 1000 if talent_raw >= 1000 else 1
                    if talent > 99: talent = 99
                    
                    info = {'id': d_id, 'name': name, 'loc': 'FAZENDA', 'slot': i, 'level': level, 'exp': exp, 'talent': talent}
                    self.search_results_map.append(info)
                    combo_options.append(f"{name} [{loc_lbl['FAZENDA']}] - {t['lvl_abbr']}{level} ({t['limite_abbr']}: {talent})".replace(",", "."))

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
                        level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                        exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                        talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]
                        talent = talent_raw // 1000 if talent_raw >= 1000 else 1
                        if talent > 99: talent = 99
                        
                        info = {'id': d_id, 'name': name, 'loc': loc, 'slot': i, 'level': level, 'exp': exp, 'talent': talent}
                        self.search_results_map.append(info)
                        combo_options.append(f"{name} [{loc_lbl[loc]}] - {t['lvl_abbr']}{level} ({t['limite_abbr']}: {talent})".replace(",", "."))

        if combo_options:
            self.combo_wish_results['values'] = combo_options
            self.combo_wish_results.current(0)
        else:
            self.combo_wish_results['values'] = []
            self.combo_wish_results.set("")
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
        key = (item.get('loc'), item.get('slot'))
        if key in active and active[key]['id'] == item['id']:
            found = active[key]
        else:
            for candidate in active.values():
                if candidate['id'] == item['id'] and candidate['name'] == item['name']:
                    found = candidate
                    break

        if found:
            item['orphaned'] = False
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

    def log(self, text, tag=None):
        self.text_area.config(state=tk.NORMAL)
        if isinstance(text, list):
            for pedaco, t in text:
                self.text_area.insert(tk.END, pedaco, t)
            self.text_area.insert(tk.END, "\n")
        else:
            self.text_area.insert(tk.END, text + "\n", tag)
        self.text_area.config(state=tk.DISABLED)
        self.text_area.yview(1.0)

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
        
    def calculate_almost_data(self, data, name_offset, name, level, talent, loc):
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
        
        return (name, level, talent, faltam, bar_str, faltam_str, loc)

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
        total_alerts, total_almost, total_lv99 = 0, 0, 0 
        
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
                level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                current_exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]

                active_digimons[("FAZENDA", i)] = {
                    'id': digimon_id, 'name': name, 'level': level, 'exp': current_exp, 'loc': "FAZENDA", 'slot': i
                }

                if talent_raw >= 1000:
                    talent = talent_raw // 1000
                    if talent > 99: talent = 99 
                    
                    if level == 99:
                        lv99_list["FAZENDA"].append((name, level, talent))
                        total_lv99 += 1
                    elif level >= talent: 
                        alerts["FAZENDA"].append((name, level, talent))
                        total_alerts += 1
                    elif level == talent - 1:
                        almost_data = self.calculate_almost_data(data, name_offset, name, level, talent, "FAZENDA")
                        almost_list.append(almost_data)
                        total_almost += 1

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

                if loc == "BOX": 
                    count_box += 1
                
                processed.add(offset)

                digimon_id = struct.unpack_from("<I", data, name_offset - 0x04)[0]
                level = struct.unpack_from("<I", data, name_offset + 0x60)[0]
                current_exp = struct.unpack_from("<I", data, name_offset + 0x64)[0]
                talent_raw = struct.unpack_from("<I", data, name_offset + 0x100)[0]

                active_digimons[(loc, i)] = {
                    'id': digimon_id, 'name': name, 'level': level, 'exp': current_exp, 'loc': loc, 'slot': i
                }

                if talent_raw >= 1000:
                    talent = talent_raw // 1000
                    if talent > 99: talent = 99 
                    
                    if level == 99:
                        lv99_list[loc].append((name, level, talent))
                        total_lv99 += 1
                    elif level >= talent: 
                        alerts[loc].append((name, level, talent))
                        total_alerts += 1
                    elif level == talent - 1:
                        almost_data = self.calculate_almost_data(data, name_offset, name, level, talent, loc)
                        almost_list.append(almost_data)
                        total_almost += 1

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

            key = (item['loc'], item['slot'])
            dig_info = None
            if key in active_digimons and active_digimons[key]['id'] == item['id']:
                dig_info = active_digimons[key]
            else:
                for candidate in active_digimons.values():
                    if candidate['id'] == item['id'] and candidate['name'] == item['name']:
                        dig_info = candidate
                        break

            if dig_info:
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

            exp_target = get_exp_needed(dig_id, target_lvl)
            exp_base = get_exp_needed(dig_id, cur_lvl)
            faltam = exp_target - cur_exp

            if faltam <= 0 or cur_lvl >= target_lvl:
                wishlist_reached.append((w_idx, w_name, cur_lvl, target_lvl, w_loc))
            else:
                prog_total = exp_target - exp_base
                prog_atual = cur_exp - exp_base
                pct = prog_atual / prog_total if prog_total > 0 else 0
                pct = max(0.0, min(1.0, pct))
                filled = int(pct * 10)
                bar_str = ("█" * filled) + ("░" * (10 - filled))
                faltam_str = f"{faltam:,}".replace(",", ".")
                wishlist_pending.append((faltam, w_idx, w_name, cur_lvl, target_lvl, bar_str, faltam_str, w_loc))

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

        # 1ª LISTA: ATINGIU O CAP
        for loc in ["PARTY", "BOX", "FAZENDA"]:
            alerts[loc].sort(key=lambda x: (x[0], x[1]))
            for name, level, talent in alerts[loc]:
                cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                self.log([
                    (f" [{loc_lbl[loc]:^7}] ", cor_tag),
                    (f"{name:<18} ({t['lvl_abbr']} {level:02d} / {t['limite_abbr']} {talent:02d}) {txt_atingiu}", "alert")
                ])

            # Itens da Wishlist que bateram a meta são "transferidos" pra cá:
            # em branco (pra diferenciar dos nativos) e mantendo o botão de remover.
            for w_idx, w_name, w_level, w_target, w_item_loc in wishlist_reached:
                if w_item_loc != loc:
                    continue
                txt_reached = t.get("target_reached", "<-- META ATINGIDA!")
                self.text_area.config(state=tk.NORMAL)
                self.text_area.insert(tk.END, f" [{loc_lbl[loc]:^7}] ", "status")
                self.text_area.insert(tk.END, f"{w_name:<18} ({t['lvl_abbr']} {w_level:02d} / {t['target_abbr']} {w_target:02d}) {txt_reached} ", "status")
                btn_del = tk.Button(self.text_area, text=" ❌ ", command=lambda idx=w_idx: self.delete_wishlist_item(idx),
                                    bg="#444444", fg="red", font=("Consolas", 8, "bold"),
                                    relief=tk.FLAT, cursor="hand2", padx=2, pady=0)
                self.text_area.window_create(tk.END, window=btn_del)
                self.text_area.insert(tk.END, "\n")
                self.text_area.config(state=tk.DISABLED)

        # 2ª LISTA: QUASE LÁ (Expansível)
        if total_almost > 0:
            self.log("-" * 75, "status")

            def toggle_almost():
                self.show_almost = not getattr(self, 'show_almost', False)
                self.save_config()
                self.process_save(filepath, filename)

            btn_text = t["btn_hide_details"] if getattr(self, 'show_almost', False) else t["btn_show_almost"]
            btn_almost = tk.Button(self.text_area, text=btn_text, command=toggle_almost,
                            bg="#333333", fg="white", font=("Consolas", 8, "bold"),
                            relief=tk.FLAT, cursor="hand2", padx=8, pady=0)

            self.text_area.config(state=tk.NORMAL)
            self.text_area.insert(tk.END, txt_quase, "almost")
            padding = max(2, 75 - len(txt_quase) - len(btn_text) - 2)
            self.text_area.insert(tk.END, " " * padding, "almost")
            self.text_area.window_create(tk.END, window=btn_almost)
            self.text_area.insert(tk.END, "\n\n")
            self.text_area.config(state=tk.DISABLED)

            if getattr(self, 'show_almost', False):
                almost_list.sort(key=lambda x: x[3])

                for name, level, talent, faltam_int, bar_str, faltam_str, loc in almost_list:
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                    self.log([
                        (f" [{loc_lbl[loc]:^7}] ", cor_tag),
                        (f"{name:<14} ({t['lvl_abbr']} {level:02d} / {t['limite_abbr']} {talent:02d}) [{bar_str}] {t['faltam_abbr']} {faltam_str:>7} EXP", "status")
                    ])

        self.log("-" * 75, "status")
        
        # RESUMO
        self.log([
            (f" {loc_lbl['BOX']}: {count_box}/999 ", "loc_box"),
            ("   |   ", "status"),
            (f" {loc_lbl['FAZENDA']}: {count_farm}/30 \n", "loc_fazenda")
        ])
        
        self.log("-" * 75, "status")
        self.log(t["lbl_summary"], "status")
        
        if total_alerts == 0 and total_almost == 0 and total_lv99 == 0:
            self.log(t["msg_all_normal"], "status")
        else:
            if total_alerts > 0:
                self.log(f"    -> {total_alerts}{t['summary_alerts']}", "alert")
            if total_almost > 0:
                self.log(f"    -> {total_almost}{t['summary_almost']}", "almost")
            if total_lv99 > 0:
                self.log(f"    -> {total_lv99}{t['summary_lv99']}", "loc_fazenda") 

        self.log("=" * 75, "header")
        
        # 3ª LISTA: LEVEL 99 (Expansível)
        if total_lv99 > 0:
            def toggle_lv99():
                self.show_lv99 = not getattr(self, 'show_lv99', False)
                self.save_config()
                self.process_save(filepath, filename)

            btn_text = t["btn_hide_details"] if getattr(self, 'show_lv99', False) else t["btn_show_lv99"]
            btn = tk.Button(self.text_area, text=btn_text, command=toggle_lv99, 
                            bg="#333333", fg="white", font=("Consolas", 9, "bold"), 
                            relief=tk.FLAT, cursor="hand2", padx=10, pady=2)
            
            self.text_area.config(state=tk.NORMAL)
            self.text_area.window_create(tk.END, window=btn)
            self.text_area.insert(tk.END, "\n\n")
            self.text_area.config(state=tk.DISABLED)

            if getattr(self, 'show_lv99', False):
                self.log(t["lbl_lv99_title"], "loc_fazenda")
                self.log("", "status")
                for loc in ["PARTY", "BOX", "FAZENDA"]:
                    lv99_list[loc].sort(key=lambda x: (x[0], x[1]))
                    for name, level, talent in lv99_list[loc]:
                        cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]
                        self.log([
                            (f" [{loc_lbl[loc]:^7}] ", cor_tag),
                            (f"{name:<18} ({t['lvl_abbr']} {level:02d} / {t['limite_abbr']} {talent:02d})", "status")
                        ])
                self.log("=" * 75, "header")

        # ==========================================
        # 4ª LISTA: WISHLIST / TARGET TRACKER
        # (só quem ainda não bateu a meta; quem bateu já foi
        #  transferido pra 1ª lista lá em cima)
        # ==========================================
        if wishlist_pending:
            wishlist_pending.sort(key=lambda x: x[0])  # Ordena da menor EXP para a maior

            self.log("-" * 75, "status")

            wishlist_title_text = t.get("wishlist_title", " 🎯 WISHLIST / METAS DE EVOLUÇÃO:")

            def toggle_wishlist():
                self.show_wishlist = not getattr(self, 'show_wishlist', False)
                self.save_config()
                self.process_save(filepath, filename)

            btn_text = t["btn_hide_details"] if getattr(self, 'show_wishlist', False) else t["btn_show_wishlist"]
            btn_wishlist = tk.Button(self.text_area, text=btn_text, command=toggle_wishlist,
                            bg="#333333", fg="white", font=("Consolas", 8, "bold"),
                            relief=tk.FLAT, cursor="hand2", padx=8, pady=0)

            self.text_area.config(state=tk.NORMAL)
            self.text_area.insert(tk.END, wishlist_title_text, "header")
            padding = max(2, 75 - len(wishlist_title_text) - len(btn_text) - 2)
            self.text_area.insert(tk.END, " " * padding, "header")
            self.text_area.window_create(tk.END, window=btn_wishlist)
            self.text_area.insert(tk.END, "\n\n")
            self.text_area.config(state=tk.DISABLED)

            if getattr(self, 'show_wishlist', False):
                for faltam_int, w_idx, name, level, target_lvl, bar_str, faltam_str, loc in wishlist_pending:
                    cor_tag = {"PARTY": "loc_party", "BOX": "loc_box", "FAZENDA": "loc_fazenda"}[loc]

                    self.text_area.config(state=tk.NORMAL)
                    self.text_area.insert(tk.END, f" [{loc_lbl[loc]:^7}] ", cor_tag)

                    msg = f"{name:<14} ({t['lvl_abbr']} {level:02d} / {t['target_abbr']} {target_lvl:02d}) [{bar_str}] {t['faltam_abbr']} {faltam_str:>7} EXP "
                    self.text_area.insert(tk.END, msg, "status")

                    # Botão Deletar Inline [X]
                    btn_del = tk.Button(self.text_area, text=" ❌ ", command=lambda idx=w_idx: self.delete_wishlist_item(idx),
                                        bg="#444444", fg="red", font=("Consolas", 8, "bold"), 
                                        relief=tk.FLAT, cursor="hand2", padx=2, pady=0)
                    
                    self.text_area.window_create(tk.END, window=btn_del)
                    self.text_area.insert(tk.END, "\n")
                    self.text_area.config(state=tk.DISABLED)

                self.log("=" * 75, "header")

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