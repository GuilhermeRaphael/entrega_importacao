# Importador Goalfy (Python)

## App portátil para a equipe (Windows 64 bits)

Distribua o ZIP ImportadorGoalfy-Windows.zip. Cada pessoa deve extrair a pasta
completa e abrir ImportadorGoalfy.exe. Não é necessário instalar Python.
A pasta _internal precisa permanecer ao lado do executável.

O config.json fica ao lado do executável e mantém as configurações entre
execuções. O pacote conserva cadastros de matrizes, vendedores e gestores,
mas remove os caminhos pessoais: cada usuário configura seus arquivos em
Configurações > Matrizes e a pasta de saída em Geral. Não inclui leads nem
histórico de envios. Veja GUIA_EQUIPE.txt para uso e atualização.

Para gerar novamente em Windows, com Python e as dependências instaladas:

```powershell
python -m pip install -r requirements.txt -r requirements-build.txt
.\build_windows.ps1
```

O script gera dist/ImportadorGoalfy e dist/ImportadorGoalfy-Windows.zip.
Outlook Desktop configurado continua necessário para envio de e-mails.

## Pasta de saída e geração

Em Configurações > Geral use Selecionar pasta de importação. O app valida
permissão de escrita antes de distribuir leads e usa o valor atual do campo
mesmo que não tenha sido salvo previamente. Não há troca silenciosa de pasta.
Se o salvamento do arquivo falhar após a distribuição, mantenha o app aberto,
corrija a pasta e clique em REPETIR SALVAMENTO DA IMPORTAÇÃO: a mesma remessa
é exportada sem selecionar outros leads. Essa recuperação fica nesta sessão.

A geração abre e salva cada matriz uma vez por remessa e indexa as linhas
elegíveis por modelo. A ordem da fila e a ordem original das linhas são mantidas.
O arquivo de importação geral continua sendo criado, com nome único por geração.

## Gestores e relatórios consolidados

Cadastre cada pessoa uma vez em Configurações > Gestores. Em Adicionar/Editar,
preencha nome e e-mail e atribua vários produtos na tabela. Cada produto possui
TIPO (GESTOR/SUBGESTOR), PADRÃO e ATIVO próprios. Use Novo produto, Atribuir /
Atualizar produto e Remover produto; SALVAR persiste o cadastro completo.
Os registros antigos são agrupados por nome/e-mail e seus vínculos preservados.
Os produtos devem corresponder ao campo PRODUTO dos leads da remessa.

ENVIAR E-MAIL P/ GESTORES exibe apenas vínculos ativos dos produtos entregues.
PADRÃO = SIM vem selecionado. Cada pessoa recebe um único e-mail com um único
Excel consolidando todos os vendedores e empresas dos produtos selecionados
para ela. Os demais produtos não são incluídos. Pessoas com a mesma seleção
de produtos reutilizam o mesmo arquivo. O relatório mantém cliente, telefone,
CPF/CNPJ, e-mail, datas, vendedor, produto, empresa, modelo e bonificação quando
presente. Os envios aos vendedores continuam usando seus grupos atuais.

CANCELAR não envia mensagens. Falhas individuais e pessoas sem e-mail são
registradas e apresentadas no resumo final. Outlook Desktop configurado é
necessário. Não são usadas planilhas para configurar gestores.

Testes isolados: python -B -m unittest test_gestores -v

## Atualizações de entrega

Os anexos de gestores e vendedores têm tabelas nativas do Excel, filtros,
linhas alternadas e todas as bordas. O título dos vendedores possui contorno
espesso. O título dos gestores é RELATÓRIO GERAL - ENTREGA DE LEADS, com data
DD/MM/YYYY HH:MM. A seleção de gestores usa o mesmo tema do app, com checkboxes,
Selecionar Todos e Limpar Seleção.

A remessa preserva matrizOrigem, o nome usado no cadastro de matrizes, permitindo
vínculos como CELULAR INRI mesmo quando o produto nos leads é CELULAR.
O campo PRODUTO e o payload para o Goalfy não são alterados. Arquivos antigos
sem essa coluna continuam válidos para reenvio ao Goalfy.

TRAVAR IMPORTAÇÃO REPETIDA vem ativado e bloqueia a segunda geração da mesma
fila nesta sessão. NOVA REMESSA limpa a fila para iniciar outra entrega.
O salvamento pendente continua podendo ser repetido sem selecionar novos leads.

SELECIONAR PLANILHA SALVA NA MÁQUINA permite escolher um .xlsx com a aba IMPORTAÇÃO,
revisar e selecionar os leads e confirmar o envio. Não altera matrizes nem
a remessa atual. Escolha somente os leads que precisam de reenvio para evitar
duplicar os que o Goalfy já aceitou. Erros aparecem no resumo final.

A gravação usa menor compressão ZIP e a exportação usa escrita em streaming.
Linhas vazias apenas formatadas não são percorridas ao selecionar os leads.
O app mostra a etapa atual (abertura, salvamento de matriz ou exportação).

Testes: python -B -m unittest test_gestores test_novas_funcionalidades -v

## Proteção no envio, digitação e layout de referência

A trava de repetição vale exclusivamente para ENVIAR PARA O GOALFY, que
consulta historico_envios_goalfy.json. Envios confirmados ficam bloqueados
inclusive após reiniciar o app; linhas que falharam podem ser tentadas novamente.
O histórico começa a registrar a partir da versão que introduziu essa proteção.

SELECIONAR PLANILHA SALVA NA MÁQUINA permite reenviar o mesmo arquivo quantas vezes
forem necessárias. Essa opção exibe todos os leads, permite selecionar quais
reenviar e não consulta nem modifica o histórico da trava do botão principal.

Produto, modelo e vendedor aceitam digitação e mostram sugestões. Use mouse,
setas e Enter para escolher uma opção cadastrada. VOLTAR PRODUTO / MODELO
recupera a seleção do último bloco adicionado, recalculando seu saldo.

Os anexos seguem o modelo informado: título branco com texto preto,
cabeçalho azul-escuro com texto branco, células brancas centralizadas,
bordas pretas e filtros. A tabela nativa não aplica cores alternadas.
Não há índice numérico no nome do anexo do gestor; os cabeçalhos de linhas
(1, 2, 3...) ficam ocultos na visualização do relatório.

A leitura leve da matriz mantém apenas linhas elegíveis e atualiza diretamente
as quatro células de entrega no XML do .xlsx. As demais partes do arquivo ficam
intactas. O arquivo atualizado é preparado temporariamente e substituído apenas
após terminar; alterações externas durante a distribuição são detectadas.

Testes completos:
python -B -m unittest test_gestores test_novas_funcionalidades test_protecao_otimizacao -v

## Correções da versão portátil

O caminho físico das matrizes conserva a grafia original do nome. A conversão
para minúsculas é usada somente em comparações, nunca para substituir o arquivo.
O botão de reenvio é roxo e se chama SELECIONAR PLANILHA SALVA NA MÁQUINA.
Testes da distribuição: python -B -m unittest test_distribuicao_windows -v
