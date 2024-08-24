# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""


    
import pandas as pd
import requests
from io import BytesIO
from collections import defaultdict
from colorama import Fore, Style

class ExcelDownloader:
    def __init__(self, file_id, cid="a148695a12b44c7c", authkey="!ABQNNAvVchCGrQs"):
        self.file_id = file_id
        self.cid = cid
        self.authkey = authkey
        self.download_url = self._generate_download_url()

    def _generate_download_url(self):
        return f"https://onedrive.live.com/download?cid={self.cid}&resid={self.file_id}&authkey={self.authkey}"

    def download_file(self):
        response = requests.get(self.download_url)
        response.raise_for_status()
        return BytesIO(response.content)

class ExcelProcessor:
    def __init__(self, file_content):
        self.file_content = file_content
        self.df = None

    def load_data(self, skiprows=3, usecols="A:I"):
        self.df = pd.read_excel(
            self.file_content,
            skiprows=skiprows,
            usecols=usecols
        )
        self.df.fillna(0, inplace=True)

    def get_collaborators_and_filieres(self):
        if self.df is None:
            print("Data has not been loaded yet.")
            return None
        return self.df[['Nom des Collaborateurs', 'Filiere/Carriere ']].to_dict(orient='records')

    def afficher_collaborateur_par_identifiant(self, identifiant):
        collaborateurs = self.df[self.df['Identifiant'] == identifiant]
        if not collaborateurs.empty:
            for index, row in collaborateurs.iterrows():
                print(f"{Fore.YELLOW}Collaborateur {index + 1}:{Style.RESET_ALL}")
                for column_name, value in row.items():
                    print(f"{Fore.GREEN}{column_name}: {value}{Style.RESET_ALL}")
                print()
            return collaborateurs
        else:
            print(f"{Fore.RED}Identifiant invalide. Veuillez entrer un identifiant valide.{Style.RESET_ALL}")
            return None

    def afficher_tous_les_collaborateurs(self):
        if self.df.empty:
            print(f"{Fore.RED}Aucune donnée à afficher.{Style.RESET_ALL}")
            return None
        
        result = []
        for index, row in self.df.iterrows():
            result.append(f"{row['Nom des Collaborateurs']} - {row['Filiere/Carriere ']}")
            print(result[-1])
        return result

    def afficher_collaborateurs_par_filiere(self, filiere):
        collaborateurs_filiere = self.df[self.df['Filiere/Carriere '] == filiere]
        if not collaborateurs_filiere.empty:
            result = collaborateurs_filiere['Nom des Collaborateurs'].tolist()
            for nom in result:
                print(f"{Fore.GREEN}{nom}{Style.RESET_ALL}")
            return result
        else:
            print(f"{Fore.RED}Aucun collaborateur trouvé pour la filière : {filiere}{Style.RESET_ALL}")
            return None

    def lister_collaborateurs_par_filiere(self):
        if self.df.empty:
            print(f"{Fore.RED}Aucune donnée à afficher.{Style.RESET_ALL}")
            return None
        
        filiere_dict = defaultdict(list)
        for index, row in self.df.iterrows():
            filiere_dict[row['Filiere/Carriere ']].append(row['Nom des Collaborateurs'])
        
        for filiere, noms in filiere_dict.items():
            print(f"{Fore.YELLOW}Filière/Carrière: {filiere}{Style.RESET_ALL}")
            for nom in noms:
                print(f"{Fore.GREEN}  - {nom}{Style.RESET_ALL}")
            print()
        
        return dict(filiere_dict)

    def speichern(self, output_file_path=None):
        """Enregistre le DataFrame actuel dans un fichier Excel."""
        if output_file_path is None:
            output_file_path = 'collaborateurs_modifies.xlsx'  # Nom de fichier par défaut
        try:
            self.df.to_excel(output_file_path, index=False)
            print(f"{Fore.GREEN}Fichier enregistré avec succès: {output_file_path}{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}Erreur lors de l'enregistrement du fichier: {e}{Style.RESET_ALL}")


# Utilisation des classes
file_id = "A148695A12B44C7C!16281"
downloader = ExcelDownloader(file_id)
file_content = downloader.download_file()

processor = ExcelProcessor(file_content)
processor.load_data()

# Exemple d'appel des méthodes
identifiant = '0BO2021'
collaborateur = processor.afficher_collaborateur_par_identifiant(identifiant)

# Appel de la méthode pour lister tous les collaborateurs par Filière/Carrière
filiere_dict = processor.lister_collaborateurs_par_filiere()
