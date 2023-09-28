package swing;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;

import javax.swing.BorderFactory;
import javax.swing.JProgressBar;

@SuppressWarnings("serial")

public class LiquidProgress extends JProgressBar {
	private final LiquidProgressUI UI;

	public int getBorderSize() {
		return borderSize;
	}

	public void setBorderSize(int borderSize) {
		this.borderSize = borderSize;
	}

	public int getSpaceSize() {
		return spaceSize;
	}

	public void setSpaceSize(int spaceSize) {
		this.spaceSize = spaceSize;
	}

	public Color getAnimateColor() {
		return animateColor;
	}

	public void setAnimateColor(Color animateColor) {
		this.animateColor = animateColor;
	}

	public Color getBorderColor() {
		return borderColor;
	}

	public void setBorderColoer(Color borderColor) {
		this.borderColor = borderColor;
	}

	public LiquidProgressUI getUI() {
		return UI;
	}

	int borderSize = 5;
	int spaceSize = 5;
	Color animateColor = new Color(125, 216, 255);
	Color borderColor = new Color(0, 178, 255);

	public LiquidProgress() {
		UI = new LiquidProgressUI(this);
        setOpaque(false);
        setFont(new Font(getFont().getFamily(), Font.BOLD, 20));
        setPreferredSize(new Dimension(100, 100));
        // Defina a cor da borda como transparente para desativá-la
        setBorder(BorderFactory.createLineBorder(new Color(0, 0, 0, 0))); // Borda transparente

        setBackground(Color.WHITE);
        setForeground(new Color(0, 178, 255));
        setUI(UI);
        setStringPainted(true);
	}
	
	public void startAnimation() {
		UI.start();
	}
	public void stopAnimation() {
		UI.stop();
	}
}
