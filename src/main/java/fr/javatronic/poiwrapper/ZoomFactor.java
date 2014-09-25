package fr.javatronic.poiwrapper;

/**
 * ZoomFactor -
 *
 * @author Sébastien Lesaint
 */
public class ZoomFactor {

  public static final ZoomFactor ZOOM_70 = new ZoomFactor(7, 10);

  private final int numerator;
  private final int denominator;

  public ZoomFactor(int numerator, int denominator) {
    this.numerator = numerator;
    this.denominator = denominator;
  }

  public int getNumerator() {
    return numerator;
  }

  public int getDenominator() {
    return denominator;
  }
}
