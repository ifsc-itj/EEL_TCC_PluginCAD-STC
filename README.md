# Plugin-AutoCAD-ExFil
Plugin para AutoCAD que integra um quadro de cargas do excel e desenha automaticamente o seu diagrama unifilar pelo AutoCAD.

## Tutoriais
### CONFIGURAR DEBUG
- Para configurar o código para debug, acesse a aba debug (depurar) e propriedades de deburação, indique a localização do programa a ser executado, conforme demonstra a imagem abaixo:

Vídeo demonstrativo -> [Vídeo](https://ifscedubr-my.sharepoint.com/:v:/g/personal/gustavo_hsz_aluno_ifsc_edu_br/EVUvF9es8y5OmbQ2LK2BSwYBfPlkswccbiHRlCd0WrTZ4g?e=d2eJHD).

![Debug](https://github.com/gustavohsz/Plugin-AutoCAD-ExFil/assets/95059305/acbe4677-0b9a-4972-b6ce-639d62d3f4df)

### EXECUÇÃO DO PROGRAMA
- Após configurado o debug, basta executar o código (F5). Localizar a pasta debug do código (conforme demonstra a imagem e vídeo) e copiar caminho da pasta. No AutoCAD ao iniciar um novo desenho digite o comando "NETLOAD" e cole o caminho da pasta do debug, dentro da pasta abra o arquivo "ExcelToAutoCAD.dll" e selecione a opção "Always Load". Feito isso o programa está carregado. Na sequência digite o comando "EXCELTOAUTOCAD" e o programa deve iniciar.
  
Vídeo demonstrativo -> [Vídeo](https://ifscedubr-my.sharepoint.com/:v:/g/personal/gustavo_hsz_aluno_ifsc_edu_br/EREJw5r8dDVGmS7ltSffn2EBkh0mLrV6cGgvU01nJGMHUg?e=1Z2Oh3).

![Dll](https://github.com/gustavohsz/Plugin-AutoCAD-ExFil/assets/95059305/15321337-0d56-4163-a1e4-880553ddc11c)


##

Referência documentação Autodesk -> [Link](https://help.autodesk.com/view/OARX/2023/ENU/?guid=GUID-7E64FDE7-C818-4566-ADF8-C40D50D91E32).

Ribbon Tab >> https://forums.autodesk.com/t5/net/net-c-ribbon-tab/m-p/3921896#M34971

##
### SITUAÇÃO ATUAL DO PROJETO:
> FUNÇÕES DESENVOLVIDAS
- [x] Planilha padronizada;
- [x] Acesso Excel
- [x] Acesso AutoCAD
- [x] Inserção de blocos por código
> LEITURA DE DADOS
- [x] Numeração de circuitos
- [x] Descrição dos circuitos
- [x] Fases
- [x] Disjuntores

### EM DESENVOLVIMENTO

GRUPOS DE CIRCUITOS COM INTERRUPTORES (DISJUNTORES) DR. 
- Verificar quais disjuntores formam cada grupo de DR, acredito que a maneira mais fácil de fazer isso é verificar primeiro os grupos de DR no código, uma maneira seria usando índices, por exemplo, o grupo de índice 0 (zero) não pertence a nenhum grupo de DRs, a partir do índice 1 (um) começa a contabilizar o grupo de DRs. 



### Pastas de estudos
- DrawObjects - Operações básicas para desenhar linhas, circulos, polilinhas, etc.
- Manipulators - Operações de copy, move, erase, scale, rotate, mirror, etc.
- Dictionaries - Operações com Layers, LineTypes e TextStyles.
- UserInputFunctions - Operações com entradas do usuário.

## User Input Functions:
<details>

- **GetString**: The GetString method prompts the user for the input of a string at the command prompt.
- **GetPoint**: The GetPoint method pormpts the user to specify a point at the command prompt.
- **GetKeyWords**: Prompts the user for input of a jeyword at the command prompt.
- **GetDistance**: Calculates the distance between two points picked by the user.

</details>

## Selection Sets:
<details>

- SelectionOnScreen (GetSelection) - prompts the user to select objets on screen
- SelectWindow - Selects all objects completely inside a rectangle defined by two points
- SelectCrossingWindow - Select objects within and crossing area defined by two points
- SelectFence - Selects all objects crossing a selection fence. Fence selection is similar to crossing polygon selection except that the fence is not closed, and a fence can cross itself
        
- PickFirstSelection 
        
        Conditions to use PickFirst: 
        - PICKFIRST  system variable = 1
        - "UsePickSet" command flag must be defined with the command that use the PickFirst selection set
        - Call the "SelectImplied" method to obtain the PickFirst selection set
        - To clear the current PickFirst selection set use the "SetImpliedSelection" method

</details>


## Hierarquia de objetos do AutoCAD

![AutoCAD Object Hierarchy](https://help.autodesk.com/cloudhelp/2023/ENU/OARX-DevGuide-Managed/images/GUID-1AA8F78F-DF90-4AA4-A975-A06FBF65231C.png)

Referência documentação Autodesk -> [Link](https://help.autodesk.com/view/OARX/2023/ENU/?guid=GUID-7E64FDE7-C818-4566-ADF8-C40D50D91E32).


## Document, Database e Editor:
<details>

Os objetos "Document", "Database" e "Editor" são todos parte do modelo de objetos do AutoCAD .NET API e são utilizados para interagir com o desenho aberto no AutoCAD.

O objeto "Document" é utilizado para acessar o desenho ativo (ou seja, o desenho que está sendo exibido no momento) no AutoCAD. Ele fornece acesso a objetos de desenho, como layers, textos, tabelas, blocos e outros elementos.

O objeto "Database" é uma classe que representa o banco de dados do desenho ativo. Ele contém as definições e geometrias de todas as entidades do desenho. A partir do objeto "Database", é possível acessar e manipular a estrutura do desenho, como tabelas de estilos, tipos de linha, camadas e outros elementos.

O objeto "Editor" é utilizado para interagir com a janela de comando do AutoCAD. Ele permite que o programa C# exiba mensagens e solicitações na janela de comando e obtenha entrada do usuário a partir dela.

Em resumo, esses objetos são utilizados em conjunto para interagir com o desenho e a interface do usuário do AutoCAD a partir do código C#.

</details>

## CommandMethod
<details>
A anotação [CommandMethod] indica que este método é um comando a ser registrado no AutoCAD.

Exemplo:
```ruby

[CommandMethod("Hello")]

```
</details>

## Transaction
<details>

O objeto "Transaction" é responsável por controlar as alterações feitas no desenho durante a execução do código.

"Transaction" é a classe que representa uma transação em um desenho AutoCAD, ou seja, uma unidade de operação que pode ser revertida. Ela é criada a partir do objeto "TransactionManager" da base de dados do desenho.

"db.TransactionManager.StartTransaction()" inicia uma nova transação no banco de dados do desenho, retornando uma instância de um objeto "Transaction".

Exemplo:
```ruby

using (Transaction trans = db.TransactionManager.StartTransaction())

```

</details>

## BlockTable

<details>

"BlockTable" é a classe que representa a tabela de blocos do desenho. Ela contém uma lista de todos os blocos definidos no desenho.

"trans.GetObject" é um método que obtém um objeto a partir de seu identificador. Neste caso, estamos obtendo a tabela de blocos a partir do ID da tabela de blocos no banco de dados do desenho.

"db.BlockTableId" é o ID da tabela de blocos no banco de dados do desenho.

"OpenMode.ForRead" é um modo de acesso que permite ler o objeto, mas não modificá-lo.

"as BlockTable" converte o objeto retornado para o tipo "BlockTable".

Exemplo:

```ruby

BlockTable bt;
bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

```

</details>

## BlockTableRecord
<details>

"BlockTableRecord" é a classe que representa um registro de tabela de bloco. Neste caso, estamos obtendo o registro de bloco do espaço do modelo.

"bt[BlockTableRecord.ModelSpace]" acessa o registro de bloco do espaço do modelo na tabela de blocos.

"as BlockTableRecord" converte o objeto retornado para o tipo "BlockTableRecord".

```ruby

BlockTableRecord btr;
btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

```

</details>


## DBObjectCollection, Entity, Matrix3d.Displacement
<details>

### DBObjectCollection 

DBObjectCollection é uma classe do .NET Framework que representa uma coleção de objetos do AutoCAD que herda da classe System.Collections.CollectionBase. Ela é usada para armazenar uma coleção de objetos do tipo DBObject, que é a classe base para muitos objetos do AutoCAD, como Entity, BlockTableRecord, BlockReference, Dimension, etc.

Uma instância de DBObjectCollection pode ser usada para armazenar vários objetos do AutoCAD em uma única coleção para que eles possam ser facilmente manipulados juntos. Ela fornece vários métodos úteis para adicionar, remover e manipular objetos em sua coleção.

### Entity

Entity é uma classe do .NET Framework que representa um objeto gráfico no desenho do AutoCAD. Ela é a classe base para muitos tipos de objetos gráficos, como linhas, arcos, círculos, polilinhas, texto, blocos, entre outros.

As instâncias de Entity contêm informações sobre a geometria do objeto, como suas coordenadas, tamanho, forma e cor. Elas também podem ter propriedades adicionais específicas para cada tipo de objeto, como largura de linha, raio, ângulo, entre outras.

As entidades são criadas dentro de um BlockTableRecord (também conhecido como espaço de modelo ou espaço de papel) e, posteriormente, podem ser adicionadas a outros objetos, como blocos e layouts, por exemplo.

### Matrix3d.Displacement

Matrix3d.Displacement é um método estático da classe Matrix3d que cria uma matriz de deslocamento. Essa matriz define um deslocamento 3D em relação aos eixos X, Y e Z.

</details>




### Estrutura básica para inserir/manipular objetos:

```ruby
[CommandMethod("Nome")]
        public static void Nome()
        {

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    //abrir o blocktable para leitura
                    BlockTable bt;
                    bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    //escreve no blocktable
                    BlockTableRecord btr;
                    btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;


                    doc.SendStringToExecute("._zoom e ", false, false, false);
                    doc.SendStringToExecute("._regen ", false, false, false);
                    trans.Commit();
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage("Error: " + ex.Message);
                    trans.Abort();

                }
            }
        }
```

### Estrutura básica manipulação de layers:

```ruby
[CommandMethod("UpdateLayer")]
        public static void UpdateLayer()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    LayerTable lyTab = trans.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;


                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage("Error: " + ex.Message);
                    trans.Abort();
                }
            }
        }
```


