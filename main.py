import datetime
import locale
import os
import webbrowser
import pandas as pd

# === CONFIGURAÇÃO DE LÍNGUA PARA EXIBIR O DIA DA SEMANA EM PORTUGUÊS ===
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    print("⚠️ Aviso: Não foi possível definir a localidade para pt_BR. Os dias da semana podem aparecer em inglês.")

def obter_data():
    hoje = datetime.date.today()
    dia_semana = hoje.strftime('%A')
    data_formatada = hoje.strftime('%d/%m/%Y')
    return f"{dia_semana.capitalize()}, {data_formatada}"

def aguardar_confirmacao():
    while True:
        resposta = input("Concluído? (s/n): ").strip().lower()
        if resposta == 's':
            break
        elif resposta == 'n':
            print("🔄 Tudo bem! Realize a atividade e depois digite 'S' para continuar.")
        else:
            print("❗ Digite apenas 's' (sim) ou 'n' (não).")

def exportar_para_excel(dados, nome_loja):
    df_dados = pd.DataFrame(list(dados.items()), columns=["Indicador", "Valor"])
    df_dados["Indicador"] = df_dados["Indicador"].str.upper()

    data_formatada = datetime.date.today().strftime("%d/%m/%Y")
    cabecalho = pd.DataFrame([[nome_loja.upper(), data_formatada]], columns=["Indicador", "Valor"])

    df_final = pd.concat([cabecalho, df_dados], ignore_index=True)

    pasta_saida = "exportados"
    os.makedirs(pasta_saida, exist_ok=True)

    nome_arquivo = f"DIARIO_{nome_loja.upper()}_{datetime.date.today().isoformat()}.xlsx"
    caminho_completo = os.path.join(pasta_saida, nome_arquivo)

    df_final.to_excel(caminho_completo, index=False, header=False)

    # Compatível com Replit para baixar o arquivo
    try:
        from replit import files
        files.download(caminho_completo)
    except ImportError:
        pass

    return caminho_completo

def coletar_dados_operacionais(loja):
    print("\n📊 Coleta de Dados Operacionais")

    indicadores_percentuais = [
        "NPS", "Improdutivo", "Abertura Tablet", "Encerramento Tablet",
        "% Recuperação de Avarias", "Ocupação do Dia", "Ocupação Acumulada"
    ]
    indicadores_inteiros = ["Reservas do Dia", "Devoluções do Dia"]
    titulos = indicadores_percentuais + indicadores_inteiros
    dados = {}

    for titulo in titulos:
        while True:
            try:
                valor = float(input(f"Digite o valor para {titulo}: ").replace(',', '.'))
                dados[titulo] = valor
                break
            except ValueError:
                print("❗ Entrada inválida. Use números. Para decimais, use ponto.")

    print("\n✅ Dados registrados com sucesso:\n")
    for titulo in titulos:
        valor = dados[titulo]
        if titulo in indicadores_percentuais:
            valor_formatado = f"{valor:.2f}".replace('.', ',') + "%"
        else:
            valor_formatado = f"{int(valor)}"
        print(f"- {titulo}: {valor_formatado}")

    caminho_arquivo = exportar_para_excel(dados, loja)
    print(f"\n📁 Dados exportados para: {caminho_arquivo}")

def gerar_planejamento(nome, loja):
    data = obter_data()
    print("\n" + "="*46)
    print(f"Planejamento Diário – {data}")
    print("="*46)
    print(f"\n👤 Gestor: {nome} | 🏢 Loja: {loja}")

    print("\n🕗 Bom dia!")
    print("\n- Conferir Escalas e Intervalos")
    print("\n- Conferir Calendários e Lembretes")
    print("\n- Conferir se há agendamentos de CPA")
    print("\n- Liberar lavagens e abastecimentos")
    print("\n- Verificar Previsão de Ocupação por Loja e Consulta de Disponibilidade")
    print("\n- Enviar dados operacionais")
    print("https://vetorzkm.movida.com.br/login.php")
    print("https://app.powerbi.com/home?experience=power-bi")

    if input("\nDeseja abrir o Power BI agora no navegador? (s/n): ").strip().lower() == 's':
        webbrowser.open("https://app.powerbi.com/home?experience=power-bi")
    if input("Deseja abrir o Vetorzkm agora no navegador? (s/n): ").strip().lower() == 's':
        webbrowser.open("https://vetorzkm.movida.com.br/login.php")

    coletar_dados_operacionais(loja)

    print("\n- Aprovações SAP LOGON e RH")
    print("https://colaborador.simpar.com.br/irj/portal")
    if input("\nDeseja abrir o SAP Logon agora no navegador? (s/n): ").strip().lower() == 's':
        webbrowser.open("https://colaborador.simpar.com.br/irj/portal")

    print("\n- Verificar pendências no Microsoft To Do, Conferir e-mails - Teams")
    print("- Reunião rápida com os líderes de equipe (15 min)")
    aguardar_confirmacao()

    print("\n🕘 Acompanhamento da Frota")
    print("\n- Verificar veículos com últimas movimentações realizadas há mais de 3 dias")
    print("\n- Conferir e cobrar o retorno dos veículos nos status Improdutivo Manutenção, Improdutivo Total e Não Operacional")
    print("\n- Conferir limpeza, abastecimento e luzes de alerta nos veículos")
    print("\n- Direcionar para a Frota veículos com manutenção pendente")
    aguardar_confirmacao()

    print("\n🕓 Estudos: Engenharia de Software")
    print("- Tema sugerido: Entrada e saída de dados em Python")
    print("- Vídeo recomendado: https://youtu.be/S9uPNppGsGo")
    if input("\nDeseja abrir o vídeo agora no navegador? (s/n): ").strip().lower() == 's':
        webbrowser.open("https://youtu.be/S9uPNppGsGo")

    print("\n🕕 Planejamento do Dia Seguinte")
    print("- Definir as prioridades para o dia seguinte")

    print("\n💡 Dicas:")
    print("Comece o dia com uma boa xícara de café ☕")
    print("Acompanhe o desempenho da equipe e faça ajustes quando necessário.")
    print("Lembre-se de que a comunicação é chave para o sucesso da equipe.")
    print("\nBoa sorte e tenha um excelente dia! 🚀")

if __name__ == "__main__":
    print("Planejador de Atividades Diárias")
    nome = input("Digite seu nome: ")
    loja = input("Digite o nome da sua loja ou unidade: ")
    gerar_planejamento(nome, loja)
