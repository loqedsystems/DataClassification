import re
import pandas as pd
import unicodedata

def remove_accents(text):
    """
    Remove acentos de uma string.
    """
    return ''.join(char for char in unicodedata.normalize('NFKD', text) if not unicodedata.combining(char))

def convert_seconds_to_hhmmss(seconds):
    """
    Converte tempo em segundos para o formato HH:MM:SS.
    """
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

def normalize_text(text):
    """
    Remove acentos e caracteres especiais de uma string, além de deixá-la em minúsculas.
    """
    text = remove_accents(text)
    return re.sub(r'[^\w\s]', '', text).lower()  # Remove caracteres especiais e normaliza para minúsculas

def normalize_process(process):
    """
    Normaliza o nome do processo: remove pontos, deixa tudo em minúsculas.
    """
    process = process.lower().replace('.', '')
    return process

def classify_social_networks(title, domain, url):
    """
    Classifica atividades relacionadas a redes sociais.
    """
    social_networks = ['facebook', 'instagram', 'twitter', 'linkedin', 'tiktok', 'snapchat', 'reddit', 'pinterest', 'tumblr', 'weibo', 'whatsapp', 'wechat']
    social_domains = ['facebook.com', 'instagram.com', 'twitter.com', 'linkedin.com', 'tiktok.com', 'snapchat.com', 'reddit.com', 'pinterest.com', 'tumblr.com', 'weibo.com', 'whatsapp.com', 'web.whatsapp.com', 'wechat.com']

    if any(network in title for network in social_networks) or \
       any(network in domain for network in social_domains) or \
       any(network in url for network in social_domains):
        classification = 'Rede Social'
        if 'facebook' in title or 'facebook.com' in domain or 'facebook.com' in url:
            subclassification = 'Facebook'
        elif 'instagram' in title or 'instagram.com' in domain or 'instagram.com' in url:
            subclassification = 'Instagram'
        elif 'twitter' in title or 'twitter.com' in domain or 'twitter.com' in url:
            subclassification = 'Twitter'
        elif 'linkedin' in title or 'linkedin.com' in domain or 'linkedin.com' in url:
            subclassification = 'LinkedIn'
        elif 'tiktok' in title or 'tiktok.com' in domain or 'tiktok.com' in url:
            subclassification = 'TikTok'
        elif 'snapchat' in title or 'snapchat.com' in domain or 'snapchat.com' in url:
            subclassification = 'Snapchat'
        elif 'reddit' in title or 'reddit.com' in domain or 'reddit.com' in url:
            subclassification = 'Reddit'
        elif 'pinterest' in title or 'pinterest.com' in domain or 'pinterest.com' in url:
            subclassification = 'Pinterest'
        elif 'tumblr' in title or 'tumblr.com' in domain or 'tumblr.com' in url:
            subclassification = 'Tumblr'
        elif 'weibo' in title or 'weibo.com' in domain or 'weibo.com' in url:
            subclassification = 'Weibo'
        elif 'whatsapp' in title or 'whatsapp.com' in domain or 'web.whatsapp.com' in url:
            subclassification = 'WhatsApp'
        elif 'wechat' in title or 'wechat.com' in domain or 'wechat.com' in url:
            subclassification = 'WeChat'
        else:
            subclassification = None
        return classification, subclassification
    return None, None

def classify_streaming_apps(title, domain, url):
    """
    Classifica atividades relacionadas a aplicativos de streaming.
    """
    streaming_apps = ['youtube', 'twitch', 'netflix', 'disney+', 'hulu', 'amazon prime', 'spotify']
    streaming_domains = ['youtube.com', 'twitch.tv', 'netflix.com', 'disneyplus.com', 'hulu.com', 'primevideo.com', 'spotify.com']

    if any(app in title for app in streaming_apps) or \
       any(app in domain for app in streaming_domains) or \
       any(app in url for app in streaming_domains):
        classification = 'Aplicativo de Streaming'
        if 'youtube' in title or 'youtube.com' in domain or 'youtube.com' in url:
            subclassification = 'YouTube'
        elif 'twitch' in title or 'twitch.tv' in domain or 'twitch.tv' in url:
            subclassification = 'Twitch'
        elif 'netflix' in title or 'netflix.com' in domain or 'netflix.com' in url:
            subclassification = 'Netflix'
        elif 'disney+' in title or 'disneyplus.com' in domain or 'disneyplus.com' in url:
            subclassification = 'Disney+'
        elif 'hulu' in title or 'hulu.com' in domain or 'hulu.com' in url:
            subclassification = 'Hulu'
        elif 'amazon prime' in title or 'primevideo.com' in domain or 'primevideo.com' in url:
            subclassification = 'Amazon Prime'
        elif 'spotify' in title or 'spotify.com' in domain or 'spotify.com' in url:
            subclassification = 'Spotify'
        else:
            subclassification = None
        return classification, subclassification
    return None, None


def classify_office_apps(title, domain, url, process):
    """
    Classifica atividades relacionadas a aplicativos de escritório.
    """
    # Normaliza o processo
    process = normalize_process(process)
    
    print(f"Processo normalizado: {process}")
    print(f"Título recebido: {title}")
    print(f"URL recebido: {url}")
    
    office_apps_processes = {
        'winwordexe': 'Microsoft Word',
        'excelexe': 'Microsoft Excel',
        'powerpointexe': 'Microsoft PowerPoint',
        'outlookexe': 'Microsoft Outlook',
        'onenoteexe': 'OneNote',
        'msteamsexe': 'Microsoft Teams',
        'powerpntexe': 'Microsoft Power Point',
        'zoomexe': 'Zoom'
        
    }

    office_apps_titles = {
        'microsoft word': 'Microsoft Word',
        'ms word': 'Microsoft Word',
        'google docs': 'Google Docs',
        'microsoft excel': 'Microsoft Excel',
        'ms excel': 'Microsoft Excel',
        'google sheets': 'Google Sheets',
        'microsoft powerpoint': 'Microsoft PowerPoint',
        'ms powerpoint': 'Microsoft PowerPoint',
        'microsoft outlook': 'Microsoft Outlook',
        'outlook': 'Microsoft Outlook',
        'onenote': 'OneNote',
        'microsoft teams': 'Microsoft Teams',
        'teams': 'Microsoft Teams'
    }

    office_apps_urls = {
        'office.com': 'Office Online',
        'docs.google.com': 'Google Docs',
        'sheets.google.com': 'Google Sheets',
        'slides.google.com': 'Google Slides',
        'outlook.office.com': 'Microsoft Outlook',
        'teams.microsoft.com': 'Microsoft Teams'
    }

    # Verifica pelo nome do processo (normalizado)
    if process in office_apps_processes:
        print(f"Classificado com base no processo: {process}")
        return 'Aplicativo de Escritório', office_apps_processes[process]

    # Verifica nos títulos
    for app_title, app_name in office_apps_titles.items():
        if app_title in title:
            print(f"Classificado com base no título: {title}")
            return 'Aplicativo de Escritório', app_name

    # Verifica nos URLs
    for app_url, app_name in office_apps_urls.items():
        if app_url in url:
            print(f"Classificado com base no URL: {url}")
            return 'Aplicativo de Escritório', app_name

    print("Não foi possível classificar este dado.")
    return None, None






def classify_internal_systems(process, title):
    """
    Classifica atividades relacionadas a sistemas internos.
    """
    internal_systems_processes = ['mstsc.exe']  
    internal_systems_titles = ['rm']

    # Verifica se ambos o processo e o título correspondem
    if process in internal_systems_processes and any(term in title for term in internal_systems_titles):
        return 'Sistema Interno', 'RM'

    return None, None



def classify_cerebro(title, domain, url):
    """
    Classifica atividades relacionadas ao sistema Cérebro.
    """
    if 'cerebro' in title or 'cerebro.df.sebrae.com.br' in domain or 'cerebro.df.sebrae.com.br' in url:
        return 'Acessos Cerebro', 'Cerebro'
    return None, None

def classify_pdf_viewer(title, url, process):
    """
    Classifica atividades relacionadas a visualizadores de PDF.
    """
    if 'pdf' in title or '.pdf' in url or 'pdf' in process:
        return 'PDF Viewer', 'PDF'
    return None, None

def classify_activity(row):
    """
    Classifica a atividade do usuário com base no título da janela, nome do processo, domínio e URL.
    """
    process = str(row['ProcessName']) if pd.notna(row['ProcessName']) else ''  # Não normalizar o processo
    title = normalize_text(str(row['WindowTitle'])) if pd.notna(row['WindowTitle']) else ''
    domain = normalize_text(str(row['Domain'])) if pd.notna(row['Domain']) else ''
    url = normalize_text(str(row['URL_Name'])) if pd.notna(row['URL_Name']) else ''

    # Verifica cada categoria de classificação
    classification, subclassification = classify_social_networks(title, domain, url)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        return row

    classification, subclassification = classify_streaming_apps(title, domain, url)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        return row

    # Chamada correta para `classify_office_apps`
    classification, subclassification = classify_office_apps(title, domain, url, process)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        return row

    classification, subclassification = classify_cerebro(title, domain, url)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        return row

    classification, subclassification = classify_pdf_viewer(title, url, process)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        return row
    
    classification, subclassification = classify_internal_systems(process, title)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        return row

    # Se não se encaixar em nenhuma das categorias
    row['Classificação'] = 'Desconhecido'
    row['SubClassificação'] = None
    return row





