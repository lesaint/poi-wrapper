package fr.javatronic.poiwrapper;

import org.apache.poi.ss.usermodel.Row;

/**
 * MergedRegionFactory -
 *
 * @author Sébastien Lesaint
 */
public final class MergedRegionFactory {
  private MergedRegionFactory() {
    // prevents instantiation
  }

  public static MergedRegion horizontal(Row row, int columnStartIndex, int columnEndIndex) {
    return new MergedRegion(row, columnStartIndex, row.getRowNum(), columnEndIndex);
  }

  public static MergedRegion vertical(Row row, int columnIndex, int rowEndIndex) {
    return new MergedRegion(row, columnIndex, rowEndIndex, columnIndex);
  }

  public static MergedRegion region(Row row, int columnStartIndex, int rowEndIndex, int columnEndIndex) {
    return new MergedRegion(row, columnStartIndex, rowEndIndex, columnEndIndex);
  }

}
