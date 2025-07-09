import pygame
import pandas as pd
import time
import textwrap
import os

# --- CONFIGURAÇÃO GERAL ---
# Alterne para False quando for usar o link
USAR_PLANILHA_LOCAL = True

# >>> PASSO 1: Configure os caminhos aqui <<<
CAMINHO_LOCAL_PLANILHA = r"D:\PainelEXE\Dados\Fila de prioridades do laboratório.xlsx"
URL_PLANILHA_ONLINE = "COLE_AQUI_A_URL_DIRETA_DA_SUA_PLANILHA_QUANDO_FOR_USAR"

# Nomes das colunas na sua planilha (ajuste se necessário)
COLUNA_PEDIDO = 'Cotação de Venda  ↑'
COLUNA_QTD = 'Quantidade Solicitada'
COLUNA_SERVICO = 'Produto / Serviço: Descrição do Produto/Serviço'
COLUNA_PV = 'Pv de Transferência'

# <<< MUDANÇA: Intervalo de verificação do arquivo e de atualização automática >>>
INTERVALO_CHECK_ARQUIVO_SEGUNDOS = 2 # Verifica se o arquivo mudou a cada 2 segundos
INTERVALO_ATUALIZACAO_SEGUNDOS = 60 # Atualização forçada a cada 60 segundos (para o modo online)


# --- INICIALIZAÇÃO E CONFIGURAÇÕES VISUAIS ---
pygame.init()

# Cores
LARANJA = (255, 102, 0)
BRANCO = (255, 255, 255)
PRETO = (0, 0, 0)
CINZA_ESCURO = (40, 40, 40)
VERMELHO_ERRO = (200, 0, 0)
BRANCO_FOSCO = (220, 220, 220)
CINZA_BOTAO_PRESSIONADO = (100, 100, 100)
CINZA_FLASH = (50, 50, 50)

# Tela com resolução dinâmica
info_tela = pygame.display.Info()
LARGURA, ALTURA = info_tela.current_w, info_tela.current_h
tela = pygame.display.set_mode((LARGURA, ALTURA), pygame.FULLSCREEN)
pygame.display.set_caption("Painel de Prioridades")

# Fontes com tamanho relativo
def get_font_size(base_size):
    return int(base_size * (ALTURA / 1080))

try:
    fonte_titulo = pygame.font.SysFont("Segoe UI", get_font_size(80), bold=True)
    fonte_bloco_titulo = pygame.font.SysFont("Segoe UI", get_font_size(40), bold=True)
    fonte_texto_pedido = pygame.font.SysFont("Segoe UI", get_font_size(38), bold=True)
    fonte_texto_servico = pygame.font.SysFont("Segoe UI", get_font_size(32))
    fonte_qtd = pygame.font.SysFont("Segoe UI", get_font_size(42), bold=True)
    fonte_erro = pygame.font.SysFont("Consolas", get_font_size(28), bold=True)
    fonte_botao = pygame.font.SysFont("Segoe UI", get_font_size(30), bold=True)
except pygame.error:
    fonte_titulo = pygame.font.SysFont("Arial", get_font_size(80), bold=True)
    fonte_bloco_titulo = pygame.font.SysFont("Arial", get_font_size(40), bold=True)
    fonte_texto_pedido = pygame.font.SysFont("Arial", get_font_size(38), bold=True)
    fonte_texto_servico = pygame.font.SysFont("Arial", get_font_size(32))
    fonte_qtd = pygame.font.SysFont("Arial", get_font_size(42), bold=True)
    fonte_erro = pygame.font.SysFont("monospace", get_font_size(28), bold=True)
    fonte_botao = pygame.font.SysFont("Arial", get_font_size(30), bold=True)

# Botão com posição e tamanho relativos
largura_botao = LARGURA * 0.11
altura_botao = ALTURA * 0.065
margem_direita_botao = LARGURA * 0.02
margem_topo_botao = ALTURA * 0.023
botao_atualizar_rect = pygame.Rect(LARGURA - largura_botao - margem_direita_botao, margem_topo_botao, largura_botao, altura_botao)


# --- FUNÇÕES PRINCIPAIS (sem alterações) ---

def ler_planilha():
    """Lê a planilha, limpa e retorna os 5 primeiros pedidos."""
    if USAR_PLANILHA_LOCAL:
        if not os.path.exists(CAMINHO_LOCAL_PLANILHA):
            raise FileNotFoundError(f"Arquivo não encontrado em: {CAMINHO_LOCAL_PLANILHA}")
        caminho_ou_url = CAMINHO_LOCAL_PLANILHA
    else:
        caminho_ou_url = URL_PLANILHA_ONLINE

    df = pd.read_excel(caminho_ou_url, engine='openpyxl')
    
    df_filtrado = df[[COLUNA_PEDIDO, COLUNA_QTD, COLUNA_SERVICO, COLUNA_PV]]
    df_filtrado.columns = ['Pedido', 'Maquinas', 'Servico', 'PV']
    
    df_limpo = df_filtrado.dropna(subset=['Pedido']).copy()
    
    df_limpo['Maquinas'] = pd.to_numeric(df_limpo['Maquinas'], errors='coerce').fillna(0).astype(int)

    return df_limpo.head(5)

def render_texto_quebrado(surface, texto, fonte, cor, rect, espaco_linha=8):
    """Quebra e renderiza o texto para caber dentro de um retângulo (rect)."""
    largura_max_pixels = rect.width
    char_largura = fonte.size('A')[0]
    if char_largura == 0: return
    max_chars_linha = max(int(largura_max_pixels / char_largura), 15)
    linhas = textwrap.wrap(str(texto), width=max_chars_linha, break_long_words=True, replace_whitespace=False)
    y = rect.y
    for linha in linhas:
        render = fonte.render(linha, True, cor)
        surface.blit(render, (rect.x, y))
        y += render.get_height() + espaco_linha

def desenhar_bloco(surface, rect, pedido_info, titulo="PRIORIDADE"):
    """Desenha um único cartão de pedido."""
    pygame.draw.rect(surface, BRANCO, rect, border_radius=15)
    margem_interna = rect.width * 0.04
    altura_cabecalho = rect.height * 0.15
    titulo_rect = pygame.Rect(rect.x + margem_interna, rect.y + margem_interna, rect.width - (margem_interna * 2), altura_cabecalho)
    pygame.draw.rect(surface, CINZA_ESCURO, titulo_rect, border_radius=10)
    texto_titulo = fonte_bloco_titulo.render(titulo, True, LARANJA)
    pos_titulo = texto_titulo.get_rect(center=titulo_rect.center)
    surface.blit(texto_titulo, pos_titulo)
    pos_y_pedido = titulo_rect.bottom + (margem_interna * 0.8)
    texto_pedido = fonte_texto_pedido.render(str(pedido_info['Pedido']), True, PRETO)
    surface.blit(texto_pedido, (rect.x + margem_interna, pos_y_pedido))
    pos_y_servico = pos_y_pedido + texto_pedido.get_height() + (margem_interna * 0.5)
    altura_area_servico = rect.bottom - pos_y_servico - (rect.height * 0.15)
    area_servico = pygame.Rect(rect.x + margem_interna, pos_y_servico, rect.width - (margem_interna * 2), altura_area_servico)
    render_texto_quebrado(surface, pedido_info['Servico'], fonte_texto_servico, PRETO, area_servico)
    texto_qtd = fonte_qtd.render(f"QTD. {pedido_info['Maquinas']}", True, PRETO)
    surface.blit(texto_qtd, (rect.x + margem_interna, rect.bottom - (rect.height * 0.12)))
    if 'teravix' not in str(pedido_info['Servico']).lower():
        pv_text = str(pedido_info.get('PV', ''))
        texto_pv_extra = fonte_texto_servico.render(f"PV: {pv_text}", True, VERMELHO_ERRO)
        pos_pv_extra = texto_pv_extra.get_rect(bottomright=(rect.right - margem_interna, rect.bottom - (margem_interna * 0.8)))
        surface.blit(texto_pv_extra, pos_pv_extra)

def desenhar_painel(df):
    """Desenha a interface principal com todos os pedidos."""
    tela.fill(CINZA_ESCURO)
    altura_cabecalho = ALTURA * 0.11
    pygame.draw.rect(tela, LARANJA, (0, 0, LARGURA, altura_cabecalho))
    titulo = fonte_titulo.render("PAINEL PRIORIDADES", True, BRANCO)
    pos_titulo = titulo.get_rect(center=(LARGURA // 2, altura_cabecalho // 2))
    tela.blit(titulo, pos_titulo)
    pygame.draw.rect(tela, BRANCO_FOSCO, botao_atualizar_rect, border_radius=10)
    texto_botao = fonte_botao.render("ATUALIZAR", True, PRETO)
    pos_texto_botao = texto_botao.get_rect(center=botao_atualizar_rect.center)
    tela.blit(texto_botao, pos_texto_botao)
    if df.empty:
        msg_aguardando = fonte_bloco_titulo.render("Aguardando novos pedidos...", True, BRANCO)
        pos_msg = msg_aguardando.get_rect(center=(LARGURA // 2, ALTURA // 2))
        tela.blit(msg_aguardando, pos_msg)
        return
    margem = LARGURA * 0.025
    espaco = LARGURA * 0.015
    y_inicial = altura_cabecalho + margem
    largura_bloco_grande = LARGURA * 0.45
    altura_bloco_grande = ALTURA - y_inicial - margem
    bloco_grande_rect = pygame.Rect(margem, y_inicial, largura_bloco_grande, altura_bloco_grande)
    desenhar_bloco(tela, bloco_grande_rect, df.iloc[0], titulo="PRIORIDADE")
    if len(df) > 1:
        x_inicial_menores = margem + largura_bloco_grande + espaco
        largura_total_menores = LARGURA - x_inicial_menores - margem
        altura_total_menores = ALTURA - y_inicial - margem
        largura_bloco_menor = (largura_total_menores - espaco) / 2
        altura_bloco_menor = (altura_total_menores - espaco) / 2
        for i in range(1, min(len(df), 5)):
            titulo_bloco_menor = f"PEDIDO {i + 1}"
            coluna = (i - 1) % 2
            linha = (i - 1) // 2
            x_pos = x_inicial_menores + coluna * (largura_bloco_menor + espaco)
            y_pos = y_inicial + linha * (altura_bloco_menor + espaco)
            bloco_menor_rect = pygame.Rect(x_pos, y_pos, largura_bloco_menor, altura_bloco_menor)
            desenhar_bloco(tela, bloco_menor_rect, df.iloc[i], titulo=titulo_bloco_menor)

def desenhar_tela_de_erro(mensagem):
    """Mostra uma mensagem de erro clara na tela."""
    tela.fill(CINZA_ESCURO)
    texto_erro_titulo = fonte_titulo.render("ERRO AO CARREGAR DADOS", True, VERMELHO_ERRO)
    pos_titulo = texto_erro_titulo.get_rect(center=(LARGURA // 2, ALTURA // 3))
    tela.blit(texto_erro_titulo, pos_titulo)
    area_msg = pygame.Rect(LARGURA * 0.05, ALTURA // 2, LARGURA * 0.9, ALTURA // 2)
    render_texto_quebrado(tela, mensagem, fonte_erro, BRANCO, area_msg)


# --- LOOP PRINCIPAL ---
executando = True
ultima_atualizacao = 0
dados_df = None
clock = pygame.time.Clock()
precisa_atualizar = True

# <<< MUDANÇA: Variáveis para controlar a verificação do arquivo >>>
timestamp_arquivo = 0
ultimo_check_arquivo = 0

while executando:
    for evento in pygame.event.get():
        if evento.type == pygame.QUIT or (evento.type == pygame.KEYDOWN and evento.key == pygame.K_ESCAPE):
            executando = False
        
        if evento.type == pygame.MOUSEBUTTONDOWN:
            if evento.button == 1 and botao_atualizar_rect.collidepoint(evento.pos):
                print("Botão de atualização clicado! Forçando a leitura da planilha.")
                precisa_atualizar = True

    agora = time.time()
    
    # <<< MUDANÇA: Lógica para detectar alteração no arquivo local >>>
    # Só executa se estiver no modo de planilha local
    if USAR_PLANILHA_LOCAL and (agora - ultimo_check_arquivo > INTERVALO_CHECK_ARQUIVO_SEGUNDOS):
        ultimo_check_arquivo = agora
        try:
            # Pega a data de modificação atual do arquivo
            mod_time_atual = os.path.getmtime(CAMINHO_LOCAL_PLANILHA)
            
            # Se a data for diferente da que temos guardada, marca para atualizar
            # A condição `timestamp_arquivo != 0` evita a atualização na primeira vez que o programa abre
            if timestamp_arquivo != 0 and mod_time_atual != timestamp_arquivo:
                print(f"[{time.strftime('%H:%M:%S')}] Detecção de mudança na planilha. Atualizando...")
                precisa_atualizar = True
            
            # Atualiza o timestamp guardado
            timestamp_arquivo = mod_time_atual

        except FileNotFoundError:
            # Se o arquivo não for encontrado, o erro será tratado no bloco principal de atualização
            pass
            
    # Define a condição para atualizar a tela
    deve_atualizar_agora = precisa_atualizar or \
                           (not USAR_PLANILHA_LOCAL and (agora - ultima_atualizacao >= INTERVALO_ATUALIZACAO_SEGUNDOS))

    if deve_atualizar_agora:
        
        start_time = time.time()

        # Feedback visual de "CARREGANDO..."
        if dados_df is not None:
             desenhar_painel(dados_df)
        else:
             tela.fill(CINZA_ESCURO)
             altura_cabecalho = ALTURA * 0.11
             pygame.draw.rect(tela, LARANJA, (0, 0, LARGURA, altura_cabecalho))
             titulo = fonte_titulo.render("PAINEL PRIORIDADES", True, BRANCO)
             pos_titulo = titulo.get_rect(center=(LARGURA // 2, altura_cabecalho // 2))
             tela.blit(titulo, pos_titulo)

        pygame.draw.rect(tela, CINZA_BOTAO_PRESSIONADO, botao_atualizar_rect, border_radius=10)
        texto_botao = fonte_botao.render("CARREGANDO...", True, BRANCO)
        pos_texto_botao = texto_botao.get_rect(center=botao_atualizar_rect.center)
        tela.blit(texto_botao, pos_texto_botao)
        pygame.display.flip()

        try:
            dados_df_novo = ler_planilha()
            dados_df = dados_df_novo 
            print(f"[{time.strftime('%H:%M:%S')}] Planilha lida com sucesso.")
            
            # <<< MUDANÇA: Atualiza o timestamp após uma leitura bem-sucedida >>>
            # Isso garante que a detecção de mudança funcione corretamente após o início
            if USAR_PLANILHA_LOCAL:
                timestamp_arquivo = os.path.getmtime(CAMINHO_LOCAL_PLANILHA)

        except Exception as e:
            print(f"[{time.strftime('%H:%M:%S')}] ERRO: {e}")
            desenhar_tela_de_erro(str(e))
            pygame.display.flip()
            pygame.time.wait(5000)
            dados_df = None

        tempo_decorrido = time.time() - start_time
        if tempo_decorrido < 1.0:
            pygame.time.wait(int((1.0 - tempo_decorrido) * 1000))

        tela.fill(CINZA_FLASH)
        pygame.display.flip()
        pygame.time.wait(100)

        ultima_atualizacao = agora
        precisa_atualizar = False # Reseta a flag após a atualização
    
    # Desenha o painel com os dados
    if dados_df is not None:
        desenhar_painel(dados_df)
    else:
        # Tela de "Aguardando" se não houver dados
        tela.fill(CINZA_ESCURO)
        altura_cabecalho = ALTURA * 0.11
        pygame.draw.rect(tela, LARANJA, (0, 0, LARGURA, altura_cabecalho))
        titulo = fonte_titulo.render("PAINEL PRIORIDADES", True, BRANCO)
        pos_titulo = titulo.get_rect(center=(LARGURA // 2, altura_cabecalho // 2))
        tela.blit(titulo, pos_titulo)
        msg_aguardando = fonte_bloco_titulo.render("Tentando carregar os dados...", True, BRANCO)
        pos_msg = msg_aguardando.get_rect(center=(LARGURA // 2, ALTURA // 2))
        tela.blit(msg_aguardando, pos_msg)
    
    pygame.display.flip()
    clock.tick(30)

pygame.quit()