package fr.javatronic.poiwrapper;

import org.apache.poi.ss.usermodel.Row;

/**
 * MergedRegion -
 *
 * @author Sébastien Lesaint
 */
public class MergedRegion {

  private static boolean debug = false;
  private static int debug_counter = 0;

  public static void enableDebug() {
    debug = true;
  }

  public static boolean isDebugEnabled() {
    return MergedRegion.debug;
  }

  public static void resetDebugCounter() {
    debug_counter = 0;
  }

  private final int counter;
  private final Row row;
  private final int colStart;
  private final int rowEnd;
  private final int colEnd;

  public MergedRegion(Row row, int colStart, int rowEnd, int colEnd) {
    this.counter = MergedRegion.debug ? debug_counter++ : 0;
    this.row = row;
    this.colStart = colStart;
    this.rowEnd = rowEnd;
    this.colEnd = colEnd;
  }

  public int getCounter() {
    return counter;
  }

  public Row getRow() {
    return row;
  }

  public int getColStart() {
    return colStart;
  }

  public int getRowEnd() {
    return rowEnd;
  }

  public int getColEnd() {
    return colEnd;
  }
}
