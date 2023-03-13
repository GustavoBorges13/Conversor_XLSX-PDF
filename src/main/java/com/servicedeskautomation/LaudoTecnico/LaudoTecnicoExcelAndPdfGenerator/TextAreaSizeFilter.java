package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.awt.Toolkit;

import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.DocumentFilter;

//esta classe que faz a limitação de caracteres
class TextAreaSizeFilter extends DocumentFilter {

	int maxCharacters;
	boolean DEBUG = false;

	public TextAreaSizeFilter(int maxChars) {
		maxCharacters = maxChars;
	}

	public void insertString(DocumentFilter.FilterBypass fb, int offs, String str, AttributeSet a)
			throws BadLocationException {
		if (DEBUG) {
			System.out.println("in DocumentSizeFilter's insertString method");
		}

		// Isto rejeita novas inserções de caracteres se
		// a string for muito grande. Outra opção seria
		// truncar a string inserida, de forma que seja
		// adicionado somente o necessário para atingir
		// o limite máximo permitido
		if ((fb.getDocument().getLength() + str.length()) <= maxCharacters) {
			super.insertString(fb, offs, str, a);
		} else {
			Toolkit.getDefaultToolkit().beep();
		}
	}

	public void replace(DocumentFilter.FilterBypass fb, int offs, int length, String str, AttributeSet a)
			throws BadLocationException {
		if (DEBUG) {
			System.out.println("in DocumentSizeFilter's replace method");
		}
		// Isto rejeita novas inserções de caracteres se
		// a string for muito grande. Outra opção seria
		// truncar a string inserida, de forma que seja
		// adicionado somente o necessário para atingir
		// o limite máximo permitido
		if ((fb.getDocument().getLength() + str.length() - length) <= maxCharacters) {
			super.replace(fb, offs, length, str, a);
		} else {
			Toolkit.getDefaultToolkit().beep();
		}
	}

}