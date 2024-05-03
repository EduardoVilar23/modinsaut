# Documentação: Ferramenta de Relatórios Fotográficos
### Versão 2.0.2 de 03/05/2024
- [Se o alerta "Risco de Segurança" for exibido](https://github.com/EduardoVilar23/modinsaut/blob/main/macroblock.md)
- [Se a mensagem "Não foi possível localizar a imagem" for exibida](https://github.com/EduardoVilar23/modinsaut/blob/main/help1.md)

---

## Introdução
Esta documentação descreve a utilização da ferramenta de relatórios fotográficos desenvolvida para auxiliar na elaboração de relatórios de revitalização de calçamento e asfalto nas ruas da cidade de Parnaíba.


## Visão Geral
A ferramenta consiste em um conjunto de módulos VBA (Visual Basic for Applications) para serem utilizados no Microsoft Excel. Os módulos permitem a inserção de imagens de buracos antes e depois da execução da revitalização, bem como a geração de um relatório fotográfico em formato de grade.


## Requisitos
- Microsoft Excel (funcionando estávelemente no Microsoft Excel 2016)
- Imagens dos buracos em formato JPEG ou JPG, nomeadas no padrão (1), (2), (3), etc.


## Utilização
### Download e Preparação de Imagens
Antes de começar, siga estas etapas para preparar suas imagens:
- **Download das Imagens**: Baixe todas as imagens necessárias para o relatório em seu computador.
- **Organização em uma Única Pasta**: Coloque todas as imagens em uma única pasta em seu computador. Evite diretórios com caminhos muito longos para evitar possíveis problemas durante a execução do programa.
- **Exclusão de Outros Arquivos**: Certifique-se de que a pasta selecionada contenha apenas as imagens referentes ao relatório nos formatos **JPEG ou JPG**.
- **Renomeação das Imagens**: Selecione todas as imagens na pasta e renomeie-as para "foto". O Windows irá numerá-las automaticamente na ordem em que foram selecionadas, seguindo o formato "foto (1)", "foto (2)", "foto (3)", e assim por diante. Essa etapa é crucial para garantir a ordem correta das imagens no relatório.
### Botões de Inserção de Imagens
Na pasta de trabalho fornecida, estão disponíveis botões de inserção de imagens nas guias correspondentes. Esses botões automatizam o processo de inserção das imagens de acordo com as especificações definidas.
- **Guia "Relatório Antes"**: Botão "Inserir Imagens"
- **Guia "Relatório"**: Botões "Inserir Imagens Antes" e "Inserir Imagens Depois"


## Limitações
- O limite máximo de imagens que podem ser inseridas é de 100.
- As imagens devem estar nomeadas no formato "foto (1)", "foto (2)", "foto (3)", e assim por diante, e devem ter as extensões **JPEG ou JPG**.
- A ferramenta pode não funcionar corretamente se as imagens estiverem em formato diferente ou se a pasta selecionada contiver outros arquivos além das imagens.


## Atualizações
- **27/04/2024**: Adicionado suporte para inserção de imagens antes e depois da execução da revitalização na guia "Relatório".
- **01/05/2024**: Adicionado suporte para imagens nos formatos JPEG e JPG.
- **03/05/2024**: Atualização do método de renomeação das imagens.
