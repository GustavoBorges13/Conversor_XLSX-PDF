package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;


import java.util.ArrayList;
import javax.swing.table.AbstractTableModel;

@SuppressWarnings("rawtypes")
public class ModeloTabela extends AbstractTableModel {
	private static final long serialVersionUID = 7738213752936430631L;
	private ArrayList linhas = null;
	private String[] colunas = null;
	
	
	public ModeloTabela(ArrayList lin, String[] col) {
		setLinhas(lin);
		setColunas(col);
	}

	public ArrayList getLinhas() {
		return linhas;
	}

	public void setLinhas(ArrayList linhas) {
		this.linhas = linhas;
	}

	public String[] getColunas() {
		return colunas;
	}

	public void setColunas(String[] colunas) {
		this.colunas = colunas;
	}
	
	//Conta as colunas e suas quantidades
	public int getColumnCount() {
		return colunas.length;
	}
	
	//Conta quantas linhas tem na tabela o nosso caso do array
	public int getRowCount() {
		return linhas.size();
	}
	
	//Responsavel por pegar o nome da coluna e retornar quantas colunas tem
	public String getColumnName(int numCol) {
		return colunas[numCol];
	}
	
	//Metodo que monta a tabela pela quantidade de linhas das colunas...	
	public Object getValueAt(int numLin, int numCol) {
		Object[] linha = (Object[])getLinhas().get(numLin);
		return linha[numCol];
	}
}
