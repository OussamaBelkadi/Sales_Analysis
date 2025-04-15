import pandas as pd

def load_data(file_path):
    """Charge les données à partir d'un fichier CSV."""
    data = pd.read_csv(file_path)
    return data

def clean_data(data):
    """Nettoie les données en supprimant les doublons et les valeurs manquantes."""
    data.drop_duplicates(inplace=True)
    data.dropna(inplace=True)
    return data

if __name__ == "__main__":
    file_path = '../data/sales_data.csv'  # Chemin vers votre fichier de données
    sales_data = load_data(file_path)
    cleaned_data = clean_data(sales_data)
    print(cleaned_data.head())
