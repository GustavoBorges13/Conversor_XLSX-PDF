# Conversor XLSX para PDF/WORD
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/GustavoBorges13/Conversor_XLSX-PDF/blob/main/LICENSE) 
[![Latest Stable Version](https://img.shields.io/badge/version-v0.2.9.11-blue)](https://github.com/GustavoBorges13/Conversor_XLSX-PDF/releases)
![GitHub last commit](https://img.shields.io/github/last-commit/GustavoBorges13/Conversor_XLSX-PDF)
<!---[![Build Status](https://app.travis-ci.com/GustavoBorges13/RunBlocker.svg?branch=main)](https://app.travis-ci.com/GustavoBorges13/RunBlocker)-->

```diff
- Em desenvolvimento!
+ Autor: Gustavo Borges Peres da Silva
# Projeto no estagio
```
# Sobre o projeto
 Esté é um Programa que realiza um espécime de automação, para tornar o trabalho desenvolvido mais rapido me levando a obter mais conhecimentos em APIs distintas que antes eu nunca tinha visto como por exemplo Apache POI (ooxml e ooxml-schemas), commons-lang, Toedter (jcalendar), jgoodies (Forms), miglayout (miglayout-layout), formdev (flatlaf) e docx4j. 
 
 O projeto foi desenvolvido utilizando MAVEN para builds e importações práticas das APIs.
 
 Sobre o programa em si, se trata de uma interface gráfica dinamica, no qual o usuario se depera com uma primeira janela para escolher a planilha em especifico para abrir, sendo que TODO o codigo foi feito especialmente para este tipo de planilha a ser trabalhado levando em considerações desde das formatações e quantidade de colunas contidas nele. Ao selecionar a planilha a mesma é transposta para uma Jtable afins de tornar a tabela editavel "como se fosse um excel". Além disso, caso o usuario selecione alguma linha da tabela, a mesma irá habilitar opções de edições e irá abrir uma janela na lateral esquerda transcrevendo os valores selecionados para a edição do mesmo. Caso o usuario esteja satisfeito, poderá selecionar a linha em especifico e exportar para um documento WORD com FORMATAÇÕES e converter logo seguinte em PDF para ser enviado ao cliente.


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
Build simples somente jar com dependencias inclusas e exec para iniciar automaticamente a aplicacao depois de compilada:
```diff
clean compile assembly:single exec:java
```
