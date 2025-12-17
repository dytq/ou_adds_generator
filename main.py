import json
from pathlib import Path
import subprocess
import random
import string


# ==============================================================
# CONFIGURATION PRINCIPALE (modulable)
# ==============================================================

"""
config = {
    "domain": "duckstore.local",
    "departements": {
        "RH": 5,
        "Compta": 3,
        "Commercial": 12,
        "Operations": 52,
        "IT": 2
    },
    "gpo_globales": [
        "GPO-Domain-Security",
        "GPO-Domain-Client-Hardening",
        "GPO-Domain-Background"
    ],
    "logiciels_msi": [
        "Chrome.msi",
        "LibreOffice.msi",
        "Antivirus.msi"
    ]
}
"""

# ============================================================
# CONFIGURATION DU FICHIER JSON
# ============================================================
USER_FILE = "users.json"       # Le fichier JSON contenant les utilisateurs
PC_FILE = "pc.json"             # Le fichier json contenant les pcs 
CONFIG_FILE = "config.json"     # Importation de la configuration json
EXPORT_PS1 = "create_users.ps1"  # Le script PowerShell généré

# ============================================================== 
# GENRATE RANDOM STRONG PASSWORD 
# ==============================================================
def generate_strong_password(length=16):
    letters = string.ascii_letters.replace('O', '').replace('I', '')
    digits = string.digits.replace('0', '').replace('1', '')
    symbols = "!@#%^&*()-_=+[]{}"

    all_chars = letters + digits + symbols

    # garantir au moins un de chaque
    password = [
        random.choice(string.ascii_lowercase),
        random.choice(string.ascii_uppercase),
        random.choice(digits),
        random.choice(symbols),
    ]

    # compléter le mot de passe
    password += [random.choice(all_chars) for _ in range(length - len(password))]

    random.shuffle(password)

    return "".join(password)



# ============================================================== 
# FONCTIONS AUTO-GENERATRICES
# ==============================================================

def generate_ou_structure(domain, departements):
    print("\n=== Génération des OU ===")
    ou_cmds = []

    ou_cmds.append(f'New-ADOrganizationalUnit -Name "PARIS" -Path "DC={domain},DC=local" -ProtectedFromAccidentalDeletion $false')

    for dpt in departements:
        ou_cmds.append(f'New-ADOrganizationalUnit -Name "{dpt}-GROUP" -Path "OU=PARIS,DC={domain},DC=local" -ProtectedFromAccidentalDeletion $false')
        ou_cmds.append(f'New-ADOrganizationalUnit -Name "{dpt}-USER" -Path "OU=PARIS,DC={domain},DC=local" -ProtectedFromAccidentalDeletion $false')
        ou_cmds.append(f'New-ADOrganizationalUnit -Name "{dpt}-PC" -Path "OU=PARIS,DC={domain},DC=local" -ProtectedFromAccidentalDeletion $false')
    return ou_cmds


def generate_groups(domain,departements):
    print("\n=== Génération des groupes AD ===")
    cmds = []

    for dpt in departements:
        # Groupe global utilisateurs
        cmds.append(f'New-ADGroup -Name "GG-{dpt}-Users" -Path "OU={dpt}-GROUP,OU=PARIS,DC={domain},DC=local" -GroupScope Global -GroupCategory Security')

        # Groupes locaux pour partages
        cmds.append(f'New-ADGroup -Name "GDL-{dpt}-Share-RW" -Path "OU={dpt}-GROUP,OU=PARIS,DC={domain},DC=local" -GroupScope DomainLocal -GroupCategory Security')
        cmds.append(f'New-ADGroup -Name "GDL-{dpt}-Share-R" -Path "OU={dpt}-GROUP,OU=PARIS,DC={domain},DC=local" -GroupScope DomainLocal -GroupCategory Security')

        # Groupe imprimante
        cmds.append(f'New-ADGroup -Name "GDL-{dpt}-Imprimante" -Path "OU={dpt}-GROUP,OU=PARIS,DC={domain},DC=local" -GroupScope DomainLocal -GroupCategory Security')

    return cmds


def generate_gpo_commands(gpo_globales, departements):
    print("\n=== Génération des GPO ===")
    cmds = []

    # GPO globales
    for gpo in gpo_globales:
        cmds.append(f'New-GPO -Name "{gpo}"')

    # GPO par département
    for dpt in departements:
        cmds.append(f'New-GPO -Name "GPO-{dpt}-DriveMap"')
        cmds.append(f'New-GPO -Name "GPO-{dpt}-Imprimante"')

    return cmds

"""
def generate_config_file(config):
    print("\n=== Sauvegarde de la configuration JSON ===")
    Path("config_AD.json").write_text(json.dumps(config, indent=4))
    print("→ Fichier config_AD.json généré.")
"""

def write_powershell_script(filename, cmds):
    print(f"=== Génération du script {filename} ===")
    Path(filename).write_text("\n".join(cmds), encoding="utf-8")
    print(f"→ Script {filename} créé.")

def generate_add_usr_gg(domain, departements):
    print("\n=== Ajout des user dans les groupes globaux utilisateurs")
    cmds = []

    for dpt in departements:
        cmds.append(f'Add-ADGroupMember -Identity "GG-{dpt}-Users" -Members (Get-ADUser -Filter * -SearchBase "OU={dpt}-USER,OU=PARIS,DC=duckstore,DC=local")')

    return cmds



# ============================================================
# 🔹 Fonction : Charger les utilisateurs depuis le JSON
# ============================================================
def load_json(json_file,name):
    with open(json_file, "r", encoding="utf-8") as file:
        data = json.load(file)
    return data[name]


# ============================================================
# 🔹 Fonction : Générer les commandes PowerShell pour AD
# ============================================================
def generate_user_commands(domain,users):
    commands = []

    for user in users:
        nom_complet = user["nom"]
        login = user["login"]
        email = user["email"]
        dpt = user["departement"]
    
        password = generate_strong_password()
        
        # OU cible
        ou_path = f'OU={dpt}-Users,OU=Paris,DC=duckstore,DC=local'

        # Commande PowerShell pour créer l'utilisateur
        cmd = (
            f'New-ADUser -Name "{nom_complet}" '
            f'-SamAccountName "{login}" '
            f'-UserPrincipalName "{login}@{domain}.local" '
            f'-EmailAddress "{email}" '
            f'-Path "{ou_path}" '
            f'-AccountPassword (ConvertTo-SecureString "{password}" -AsPlainText -Force) '
            f'-Enabled $true'
        )

        commands.append(cmd)

    return commands

def generate_pc_commands(departements):
    print("\n=== Ajout des PCs pour chaque departements")
    cmds = []
    
    for dpt in departements:
        pcs_name = departements[dpt]
        for pc_name in pcs_name:
            cmds.append(f'New-ADComputer -Name "{pc_name}" -Path "OU={dpt}-PC,OU=PARIS,DC=duckstore,DC=local"')
    
    return cmds

# ============================================================
# 🔹 Fonction : Sauvegarder les commandes dans un .ps1
# ============================================================
def save_ps1(filename, commands):
    Path(filename).write_text("\n".join(commands), encoding="utf-8")
    print(f"✔ Script PowerShell généré : {filename}")


# ==============================================================
# PROGRAMME PRINCIPAL : Génération complète
# ==============================================================

if __name__ == "__main__":
    print("=== Outil modulable de génération d'infrastructure AD ===")
    print("=== Création automatique de comptes Active Directory  ===")

    # Charger les utilisateurs
    users = load_json(USER_FILE,"users")
    pc = load_json(PC_FILE,"pc")
    config = load_json(CONFIG_FILE,"config")

    # Creation des commandes
    ou_cmds = generate_ou_structure(config["domain"], config["departements"])
    group_cmds = generate_groups(config["domain"], config["departements"])
    gpo_cmds = generate_gpo_commands(config["gpo_globales"], config["departements"])
    usr_cmds = generate_user_commands(config["domain"],users)
    pc_cmds = generate_pc_commands(config["departements"])
    add_usr_gg_cmds = generate_add_usr_gg(config["domain"], config["departements"])

    # Génération des fichiers
    # generate_config_file(config)
    write_powershell_script("01_create_USR.ps1",usr_cmds)
    write_powershell_script("02_create_OU.ps1", ou_cmds)
    write_powershell_script("03_create_Groups.ps1", group_cmds)
    write_powershell_script("04_create_GPO.ps1", gpo_cmds)
    write_powershell_script("05_create_ADD_USR.ps1", add_usr_gg_cmds)
    write_powershell_script("06_create_PC.ps1", pc_cmds)

    print("\n✔ Tous les scripts et configurations ont été générés !")
