# Conversor XLSX para PDF/WORD
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/GustavoBorges13/Conversor_XLSX-PDF/blob/main/LICENSE) 
[![Latest Stable Version](https://img.shields.io/badge/version-v0.2.9.3-blue)](https://github.com/GustavoBorges13/Conversor_XLSX-PDF/releases)
[![Latest Bug Version](https://img.shields.io/badge/version-v0.2.10.3-orange)](https://github.com/GustavoBorges13/Conversor_XLSX-PDF/releases)
[![Latest Actual-Development Version](https://img.shields.io/badge/version-v0.2.11.X-yellow)](https://github.com/GustavoBorges13/Conversor_XLSX-PDF/releases)
[![GitHub last commit](https://img.shields.io/github/last-commit/GustavoBorges13/Conversor_XLSX-PDF)](https://github.com/GustavoBorges13/Conversor_XLSX-PDF/commits/main)
<!---[![Build Status](https://app.travis-ci.com/GustavoBorges13/RunBlocker.svg?branch=main)](https://app.travis-ci.com/GustavoBorges13/RunBlocker)-->

```diff
- Finalizado e Obsoleto!
+ Autor: Gustavo Borges Peres da Silva
# Projeto no estagio
```

# Sobre o projeto
Relatorio disponivel aqui
[relatorio_projeto.pdf](https://github.com/user-attachments/files/17529859/relatorio_projeto.pdf)

# Sobre o projeto

Desenvolvemos esta aplicação para aprimorar a eficiência de nossa equipe de service desk.

Esta é uma ferramenta de automação criada com fins educacionais, visando aprofundar nosso conhecimento em várias APIs, como APACHE POI, JXL, OpenCSV e docx4j, e aprimorar nossas habilidades de programação em um ambiente profissional.

O projeto foi desenvolvido no Eclipse (versão de 2022-09) e construído com o Maven para realizar tarefas como limpeza, verificação, instalação de pacotes e builds. Isso garantiu que todas as APIs necessárias estivessem disponíveis ao alternar entre ambientes domésticos e empresariais, facilitando a colaboração por meio do GitHub, onde você pode acessar nosso perfil e o repositório deste projeto para obter mais informações.

A aplicação oferece uma interface gráfica dinâmica. Ao abri-la, o usuário pode escolher uma planilha específica para manipulação, eliminando a necessidade de usar o Excel. O código foi desenvolvido sob medida para esse tipo de planilha, considerando formatações e quantidade de colunas.

Ao selecionar a planilha, ela é exibida em uma JTable, permitindo a edição como se fosse uma planilha Excel. Além disso, ao clicar em uma linha, você terá opções de edição, e ao clicar duas vezes em uma linha, uma janela lateral abrirá para edição dos valores selecionados. O menu "Ferramentas" oferece utilidades adicionais para explorar.

Após realizar as edições desejadas, o usuário pode selecionar uma ou várias linhas e gerar um arquivo PDF. O programa automatiza a criação de um documento MS Word formatado com campos de texto da empresa, servindo como backup. Após a conclusão, o programa converterá o arquivo Word em PDF e permitirá visualizá-lo ou acessar a pasta de destino.


# Como utilizar (Tutorial em desenvolvimento)
## Tela de carregamento (SplashAnimation.java)
## Tela Inicial (Principal.java)
## Ajuda (Menu_1)
### Repositorio desde projeto (JDialog)
### Sobre E-ServiceDesk Application (JDialog)
## Ferramentas (Menu_2)
### Opções (Opcoes.java)
#### Geral (TabbedPane_panel_1)
#### Paths (TabbedPane_panel_2)
## Editar informações da planilha (EditarPlanilha.java)
## Gerar arquivo em PDF (GerarLaudoPDF.java)
### Links de referencia (JDialog)
## Atalhos (Atalhos.java)
### Tela Inicial
### Editar informações da planilha
### Gerar arquivo em PDF



# Lembretes
O programa em si foi feito utilizando jdk 20+ e precisa de jre 1.8 para rodar na maquina.
Ao transferir para outro PC o projeto raiz, certificar-se de configurar IDE para rodar em jdk 20+ e configurar o classpath com essa config:

```diff
<?xml version="1.0" encoding="UTF-8"?>
<classpath>
	<classpathentry kind="src" path="src/main/java"/>
	<classpathentry kind="src" path="resources"/>
	<classpathentry kind="con" path="org.eclipse.jdt.launching.JRE_CONTAINER/org.eclipse.jdt.internal.debug.ui.launcher.StandardVMType/JavaSE-1.8">
		<attributes>
			<attribute name="maven.pomderived" value="true"/>
		</attributes>
	</classpathentry>
	<classpathentry kind="con" path="org.eclipse.m2e.MAVEN2_CLASSPATH_CONTAINER">
		<attributes>
			<attribute name="maven.pomderived" value="true"/>
		</attributes>
	</classpathentry>
	<classpathentry kind="output" path="target/classes"/>
</classpath>
```
E também verificar as versoes de alguns recursos dentro do arquivo .pom para manter atualizado pois caso deia problema no pacote final (arquivo jar) é porque deu algum bug com windows 11 relacionado a versão de alguma biblioteca. 
Manter a linha de resource no arquivo .pom para evitar o erro de acessar as imagens no pacote de recursos:
```diff
		<resources>
			<resource>
				<directory>src/main/java</directory>
				<filtering>true</filtering>
			</resource>
			<resource>
				<directory>src/main/resource</directory>
				<filtering>true</filtering>
			</resource>
		</resources>
```
E por ultimo, a configuração de build utilizada no moven (goals).
Build completa:
```diff
clean install
clean compile install verify
```
Build simples somente jar com dependencias inclusas e exec para iniciar automaticamente a aplicacao depois de compilada (atualmente eu uso essa para criar o arquivo executavel posteriormente no Launch4j, Inno Setup, JSmooth, Exe4j, LaunchAnywhere, install4j, Excelsior ou JexePack):
```diff
clean compile assembly:single exec:java
```


# Maven Dependency Utilizadas no Projeto
Aqui está uma breve descrição de cada uma das dependências utilizadas no projeto (APIs, frameworks e utilitários).
## Relatório de Dependências (SBOM)
Para uma lista completa de dependências de software e informações sobre licenças, você pode consultar nosso [Relatório de Dependências](./.SBOM/Conversor_XLSX-PDF_GustavoBorges13_f91151e97bee266e8cdd1752ef04a09b1266cc1b.json).
Este relatório foi gerado automaticamente para garantir a transparência e o cumprimento de licenças de código aberto em nosso projeto.

## Relatório de Dependências Resumo
1. **org.apache.poi**: API - O Apache POI é uma API para criar e manipular documentos do Microsoft Office, como arquivos do Word e Excel. Ele permite que você trabalhe com esses formatos de arquivo em Java.

2. **org.apache.poi-ooxml**: API - Uma extensão do Apache POI que lida especificamente com formatos de arquivo Office Open XML, como arquivos .xlsx (Excel).

3. **org.apache.poi.ooxml-schemas**: Utilitário - Este pacote fornece as bibliotecas de esquema XML necessárias para trabalhar com arquivos Office Open XML. É uma dependência de apoio para o Apache POI.

4. **org.apache.pdfbox**: API - O Apache PDFBox é uma biblioteca Java que permite criar e manipular arquivos PDF. É uma API para trabalhar com PDFs em Java. Agradeço muito essa biblioteca pois é a melhor que eu encontrei e gratuita.

5. **org.swinglabs.pdf-renderer**: Framework - O PDF Renderer é um framework Java para renderizar arquivos PDF em componentes Swing. Ele fornece uma visualização de PDF em uma interface gráfica. Usei ele em especifico para mostrar uma visualizacao final do PDF que foi gerado na conversão para nao precisar abrir.

6. **com.documents4j-local**: Framework - O Documents4j é uma biblioteca que permite a conversão de documentos entre diferentes formatos, como Word para PDF. Esta é uma dependência para uso local do Documents4j.

7. **com.documents4j-transformer-msoffice-word**: Framework - Uma extensão do Documents4j para a transformação de documentos do Microsoft Word para outros formatos. É uma dependência para o uso de documentos do Word com o Documents4j.

8. **uk.gov.gchq.gaffer.bitmap-library**: Utilitário - Esta é uma biblioteca para trabalhar com mapas de bits. É um utilitário para manipulação de imagens em Java.

9. **commons-lang**: Utilitário - Apache Commons Lang é uma biblioteca de utilitários para fornecer funcionalidades adicionais à linguagem Java padrão. Ele oferece uma variedade de classes e métodos auxiliares.

10. **com.toedter.jcalendar**: Framework - O JCalendar é um framework para trabalhar com calendários e escolher datas em aplicações Java Swing. Não cheguei a usar ainda, estou testando ele pois tem uma visualização dinamica de calendario.

11. **com.jgoodies.jgoodies-forms**: Framework - O JGoodies Forms é um framework para criar interfaces gráficas de usuário (GUI) baseadas em Swing em Java. Ele fornece um layout de formulário flexível.

12. **com.miglayout-swing**: Framework - O MiGLayout é um layout manager para aplicações Java Swing que oferece um layout flexível e poderoso.

13. **com.formdev.flatlaf**: Framework - FlatLaf é um framework para criar interfaces de usuário em estilo "flat" (sem sombreamento) em aplicações Java Swing. Ele fornece um visual moderno para interfaces gráficas. Usei ele para os temas das janelas.
