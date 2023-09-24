package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.EventQueue;

import javax.swing.JDialog;

public class Atalhos extends JDialog {

	private static final long serialVersionUID = 7623461976504945799L;
	
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Atalhos dialog = new Atalhos();
					dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
					dialog.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the dialog.
	 */
	public Atalhos() {
		setTitle("Atalhos");
		setBounds(100, 100, 450, 300);

	}

}
