package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.util.ArrayList;
import javax.swing.table.AbstractTableModel;

@SuppressWarnings("rawtypes")
public class ModeloTabela extends AbstractTableModel {
	private static final long serialVersionUID = 7738213752936430631L;
	private ArrayList<Object[]> linhas = null;
	private ArrayList colunas = null;

	public ModeloTabela(ArrayList<Object[]> lin, ArrayList col) {
		setLinhas(lin);
		setColunas(col);
	}

	public ArrayList getLinhas() {
		return linhas;
	}

	public void setLinhas(ArrayList<Object[]> linhas) {
		this.linhas = linhas;
	}

	public ArrayList getColunas() {
		return colunas;
	}

	public void setColunas(ArrayList colunas) {
		this.colunas = colunas;
	}

	// Conta as colunas e suas quantidades
	public int getColumnCount() {
		return colunas.size();
	}

	// Conta quantas linhas tem na tabela o nosso caso do array
	public int getRowCount() {
		return linhas.size();
	}

	// Responsavel por pegar o nome da coluna e retornar quantas colunas tem
	public String getColumnName(int numCol) {
		return colunas.get(numCol).toString();
	}

	// Metodo que monta a tabela pela quantidade de linhas das colunas...
	public Object getValueAt(int numLin, int numCol) {
		Object[] linha = (Object[]) getLinhas().get(numLin);
		return linha[numCol];
	}

	public void addRow(Object[] row) {
		linhas.add(row);
		fireTableRowsInserted(linhas.size() - 1, linhas.size() - 1);
	}
	
	public void removeRow(int row) {
		this.linhas.remove(row);
		this.fireTableRowsDeleted(row, row);
	}
}