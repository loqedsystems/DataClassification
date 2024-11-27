import re
import pandas as pd
import unicodedata
import datetime



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
    Classifica atividades relacionadas a redes sociais, com WhatsApp separado.
    """
    social_networks = ['facebook', 'instagram', 'twitter', 'linkedin', 'tiktok', 'snapchat', 'reddit', 'pinterest', 'tumblr', 'weibo']
    social_domains = ['facebook.com', 'instagram.com', 'twitter.com', 'linkedin.com', 'tiktok.com', 'snapchat.com', 'reddit.com', 'pinterest.com', 'tumblr.com', 'weibo.com']

    # Normalizando os textos antes da classificação
    title = normalize_text(title)
    domain = normalize_text(domain)
    url = normalize_text(url)

    # Classificação separada para WhatsApp
    if 'whatsapp' in title or 'whatsapp.com' in domain or 'web.whatsapp.com' in url:
        return 'WhatsApp', 'WhatsApp', 'Acesso Pessoal'

    # Verifica se qualquer rede social está presente em qualquer parte do título, domínio ou URL
    if any(network in title for network in social_networks) or \
       any(network in domain for network in social_networks) or \
       any(network in url for network in social_networks):
        classification = 'Pessoais'
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
        elif 'pinterest' in title or 'pinterest' in domain or 'pinterest' in url:
            subclassification = 'Pinterest'
        elif 'tumblr' in title or 'tumblr.com' in domain or 'tumblr.com' in url:
            subclassification = 'Tumblr'
        elif 'weibo' in title or 'weibo.com' in domain or 'weibo.com' in url:
            subclassification = 'Weibo'
        else:
            subclassification = None
        return classification, subclassification, 'Acesso Pessoal'
    return None, None, 'Outros'




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

def classify_shopping_sites(title, domain, url):
    """
    Classifica atividades relacionadas a sites de compras.
    """
    shopping_sites = ['shopee', 'aliexpress', 'mercado livre', 'olx', 'amazon']
    shopping_domains = ['shopee.com', 'aliexpress.com', 'mercadolivre.com', 'olx.com', 'amazon.com']

    if any(site in title for site in shopping_sites) or \
       any(site in domain for site in shopping_domains) or \
       any(site in url for site in shopping_domains):
        classification = 'Pessoais'
        if 'shopee' in title or 'shopee.com' in domain or 'shopee.com' in url:
            subclassification = 'Shopee'
        elif 'aliexpress' in title or 'aliexpress.com' in domain or 'aliexpress.com' in url:
            subclassification = 'AliExpress'
        elif 'mercado livre' in title or 'mercadolivre.com' in domain or 'mercadolivre.com' in url:
            subclassification = 'Mercado Livre'
        elif 'olx' in title or 'olx.com' in domain or 'olx.com' in url:
            subclassification = 'OLX'
        elif 'amazon' in title or 'amazon.com' in domain or 'amazon.com' in url:
            subclassification = 'Amazon'
        else:
            subclassification = None
        return classification, subclassification
    return None, None

def classify_office_apps(title, domain, url, process):
    """
    Classifica atividades relacionadas a aplicativos de escritório.
    """
    process = normalize_process(process).lower()  # Normaliza para letras minúsculas

    # Dicionário com mapeamentos de processos, títulos e URLs para aplicativos de escritório
    office_apps = {
        'process': {
            'winwordexe': 'Microsoft Word',
            'excelexe': 'Microsoft Excel',
            'powerpointexe': 'Microsoft PowerPoint',
            'outlookexe': 'Microsoft Outlook',
            'onenoteexe': 'OneNote',
            'msteamsexe': 'Microsoft Teams',
            'zoomexe': 'Zoom'
        },
        'domain': {
            'office.com': 'Office Online',
            'docs.google.com': 'Google Docs',
            'sheets.google.com': 'Google Sheets',
            'slides.google.com': 'Google Slides',
            'outlook.office.com': 'Microsoft Outlook',
            'teams.microsoft.com': 'Microsoft Teams',
            'zoom.us': 'Zoom',
            'meet.google.com': 'Google Meet'
        },
        'url': {
            'office.com': 'Office Online',
            'docs.google.com': 'Google Docs',
            'sheets.google.com': 'Google Sheets',
            'slides.google.com': 'Google Slides',
            'outlook.office.com': 'Microsoft Outlook',
            'teams.microsoft.com': 'Microsoft Teams',
            'zoom.us': 'Zoom',
            'meet.google.com': 'Google Meet'
        },
        'title': {
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
            'teams': 'Microsoft Teams',
            'zoom meeting': 'Zoom',
            'google meet': 'Google Meet'
        }
    }

    # 1. Verifica se o processo corresponde a algum aplicativo de escritório
    for app_process, app_name in office_apps['process'].items():
        if app_process in process:
            return 'Aplicativo de Escritório', app_name

    # 2. Verifica se o domínio corresponde a algum aplicativo de escritório
    domain_lower = domain.lower()
    for app_domain, app_name in office_apps['domain'].items():
        if app_domain in domain_lower:
            return 'Aplicativo de Escritório', app_name

    # 3. Verifica se a URL contém algum termo relacionado a um aplicativo de escritório
    url_lower = url.lower()
    for app_url, app_name in office_apps['url'].items():
        if app_url in url_lower:
            return 'Aplicativo de Escritório', app_name

    # 4. Verifica se o título contém um termo relacionado a um aplicativo de escritório
    title_lower = title.lower()
    for app_title, app_name in office_apps['title'].items():
        if app_title in title_lower:
            return 'Aplicativo de Escritório', app_name

    # Se não corresponder a nenhum critério, retorna None
    return None, None



# def classify_internal_systems(process, title):
#     """
#     Classifica atividades relacionadas a sistemas internos.
#     """
#     internal_systems_processes = ['mstsc.exe']
#     internal_systems_titles = ['rm']

#     if process in internal_systems_processes and any(term in title for term in internal_systems_titles):
#         return 'Sistema Interno', 'RM'

#     return None, None

def classify_sebrae(title, domain, url,process):
    """
    Classifica atividades relacionadas aos sistemas do Sebrae, como Cérebro, RM, Outlook, e outros sistemas internos.
    """
    # Normalizando todas as strings
    title = normalize_text(title)
    domain = normalize_text(domain)
    url = normalize_text(url)

    # Definindo as palavras-chave para cada sistema interno do Sebrae
    sebrae_systems = ['cerebro', 'rm.exe', 'outlook', 'pdf']
    sebrae_domains = ['cerebro.com', 'rm.com', 'outlook.office.com', 'pdf']

    # Verificando presença de palavras-chave de sistemas internos (Cérebro, RM, Outlook, PDF) no título, domínio ou URL
    if any(system in title for system in sebrae_systems) or \
       any(system in domain for system in sebrae_systems) or \
       any(system in url for system in sebrae_domains):
        if 'cerebro' in title or 'cerebro' in domain or 'cerebro' in url:
            return 'Acessos Sebrae', 'Cérebro'
        elif 'rm' in title or 'rm' in domain or 'rm' in url:
            return 'Acessos Sebrae', 'RM'
        elif 'outlook' in title or 'outlook' in domain or 'outlook' in url:
            return 'Acessos Sebrae', 'Outlook'
        elif 'pdf' in title or '.pdf' in url:
            return 'Acessos Sebrae', 'PDF'
        else:
            return 'Acessos Sebrae', None  # Se o sistema for identificado, mas sem subcategoria clara
    return None, None

def classify_skype_activity(title, url, process):
    """
    Classifica atividades relacionadas ao uso do Skype.
    """
    if 'skype' in title or 'skype.com' in url or 'skype' in process:
        return 'Comunication', 'Skype Activity'
    return None, None


def classify_pdf_viewer(title, url, process):
    """
    Classifica atividades relacionadas a visualizadores de PDF.
    """
    if 'pdf' in title or '.pdf' in url or 'pdf' in process:
        return 'PDF Viewer', 'PDF'
    return None, None

def get_day_of_year(date_str):
    """
    Converte uma string de data (YYYY-MM-DD) para o dia do ano.
    """
    date = datetime.datetime.strptime(date_str, '%Y-%m-%d')
    return date.timetuple().tm_yday

def classify_development_apps(title, domain, url, process):
    """
    Classifica atividades relacionadas a aplicativos de desenvolvimento.
    """
    process = normalize_process(process).lower()  # Normaliza para letras minúsculas

    # Lista de mapeamentos de processos, títulos e URLs específicos para aplicativos de desenvolvimento
    dev_apps = {
        'process': {
            'vscode': 'VS Code',
            'gitexe': 'Git',
            'githubdesktopexe': 'GitHub Desktop',
            'mysqlworkbenchexe': 'MySQL Workbench',
            'sqlserverexe': 'SQL Server',
            'intellijexe': 'IntelliJ IDEA',
            'pycharmecexe': 'PyCharm',
            'eclipsecppexe': 'Eclipse',
            'sublime_textexe': 'Sublime Text',
            'postmanexe': 'Postman',
            'dockerdesktopexe': 'Docker Desktop',
            'terminalexe': 'Terminal',
            'ssmsexe': 'SQL Server Management Studio',
            'notepadpp.exe': 'Notepad++'
        },
        'domain': {
            'github.com': 'GitHub',
            'gitlab.com': 'GitLab',
            'bitbucket.org': 'Bitbucket',
            'stackoverflow.com': 'Stack Overflow'
        },
        'url': {
            'github.com': 'GitHub',
            'gitlab.com': 'GitLab',
            'bitbucket.org': 'Bitbucket',
            'stackoverflow.com': 'Stack Overflow'
        },
        'title': {
            'visual studio code': 'VS Code',
            'mysql workbench': 'MySQL Workbench',
            'sql server': 'SQL Server',
            'intellij': 'IntelliJ IDEA',
            'pycharm': 'PyCharm',
            'eclipse': 'Eclipse',
            'sublime text': 'Sublime Text',
            'postman': 'Postman',
            'docker': 'Docker',
            'terminal': 'Terminal',
            'notepad++': 'Notepad++',
            'stack overflow': 'Stack Overflow'
        }
    }

    # 1. Verifica se o processo corresponde a algum aplicativo de desenvolvimento
    for app_process, app_name in dev_apps['process'].items():
        if app_process in process:
            return 'Aplicativos de Desenvolvimento', app_name

    # 2. Verifica se o domínio corresponde a algum aplicativo de desenvolvimento
    domain_lower = domain.lower()
    for app_domain, app_name in dev_apps['domain'].items():
        if app_domain in domain_lower:
            return 'Aplicativos de Desenvolvimento', app_name

    # 3. Verifica se a URL contém algum termo relacionado a um aplicativo de desenvolvimento
    url_lower = url.lower()
    for app_url, app_name in dev_apps['url'].items():
        if app_url in url_lower:
            return 'Aplicativos de Desenvolvimento', app_name

    # 4. Verifica se o título contém um termo relacionado a um aplicativo de desenvolvimento
    title_lower = title.lower()
    for app_title, app_name in dev_apps['title'].items():
        if app_title in title_lower:
            return 'Aplicativos de Desenvolvimento', app_name

    # Se não corresponder a nenhum critério, retorna None
    return None, None





# Atualizando a função classify_activity para incluir a nova categoria de aplicativos de desenvolvimento
def classify_activity(row):
    """
    Classifica a atividade do usuário com base no título da janela, nome do processo, domínio e URL.
    Adiciona uma nova coluna para categorizar como "Acesso Pessoal", "Acesso Sebrae", ou "Outros".
    """
    process = str(row['ProcessName']) if pd.notna(row['ProcessName']) else ''
    title = normalize_text(str(row['WindowTitle'])) if pd.notna(row['WindowTitle']) else ''
    domain = normalize_text(str(row['Domain'])) if pd.notna(row['Domain']) else ''
    url = normalize_text(str(row['URL_Name'])) if pd.notna(row['URL_Name']) else ''

    # Verifica cada categoria de classificação
    classification, subclassification, tipo = classify_social_networks(title, domain, url)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        row['Tipo'] = tipo
        return row

    classification, subclassification = classify_streaming_apps(title, domain, url)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        row['Tipo'] = 'Acesso Pessoal'  # Streaming é considerado pessoal
        return row

    classification, subclassification = classify_office_apps(title, domain, url, process)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        row['Tipo'] = 'Acesso Sebrae'  
        return row

    classification, subclassification = classify_shopping_sites(title, domain, url)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        row['Tipo'] = 'Acesso Pessoal'  # Sites de compras são considerados pessoais
        return row

    classification, subclassification = classify_development_apps(title, domain, url, process)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        row['Tipo'] = 'Acesso Sebrae'  # Aplicativos de Desenvolvimento são categorizados como Outros
        return row

    classification, subclassification = classify_sebrae(title, domain, url,process)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        row['Tipo'] = 'Acesso Sebrae'  
        return row

    classification, subclassification = classify_pdf_viewer(title, url, process)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        row['Tipo'] = 'Acesso Sebrae'  
        return row
    

    classification, subclassification = classify_skype_activity(title, url, process)
    if classification:
        row['Classificação'] = classification
        row['SubClassificação'] = subclassification
        row['Tipo'] = 'Acesso Sebrae'  # Skype é considerado um aplicativo de comunicação
        return row

    # classification, subclassification = classify_internal_systems(process, title)
    # if classification:
    #     row['Classificação'] = classification
    #     row['SubClassificação'] = subclassification
    #     row['Tipo'] = 'Acesso Sebrae'  # Sistemas internos são categorizados como Outros
    #     return row

    # Se não for classificado em nenhuma categoria específica, é categorizado como "Outros"
    row['Classificação'] = 'Outros'
    row['SubClassificação'] = None
    row['Tipo'] = 'Outros'
    return row

