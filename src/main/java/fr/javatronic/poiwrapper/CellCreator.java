package fr.javatronic.poiwrapper;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.LoggerFactory;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;

public class CellCreator {
  private static final org.slf4j.Logger LOGGER = LoggerFactory.getLogger(CellCreator.class);

  public final static short BOLDWEIGHT_BOLD = 0x2bc;

  private static int threshold = Integer.MAX_VALUE;

  public static void setThreshold(int threshold) {
    CellCreator.threshold = threshold;
  }

  // common CellStyle methods
  private static void setCellStyle(@Nonnull Cell cell,
                                   @Nullable CellStyleDescriptor style, @Nullable Border border, @Nullable DataFormat dataformat) {
    if (style != null || border != null || dataformat != null) {
      cell.setCellStyle(createCellStyle(cell.getRow(), style, border, dataformat));
    }
  }

  public static CellStyle createCellStyle(@Nonnull Row row,
                                          @Nullable CellStyleDescriptor cellStyleDescriptor,
                                          @Nullable Border border, @Nullable DataFormat dataformat) {
    Workbook wb = row.getSheet().getWorkbook();

    CellStyle style = wb.createCellStyle();
    setFont(wb, style, cellStyleDescriptor);
    setBGColor(style, cellStyleDescriptor);
    setAlignements(style, cellStyleDescriptor);
    setBorders(style, border);
    if (dataformat != null) {
      style.setDataFormat(dataformat.getFormat());
    }
    if (cellStyleDescriptor != null) {
      style.setWrapText(cellStyleDescriptor.isWrapText());
    }
    return style;
  }

  private static void setAlignements(@Nonnull CellStyle style, @Nullable CellStyleDescriptor descriptor) {
    if (descriptor == null) {
      return;
    }

    style.setVerticalAlignment(descriptor.getAlignV());
    if (descriptor.getAlignH() != null) {
      style.setAlignment(descriptor.getAlignH());
    }
  }

  private static void setFont(@Nonnull Workbook wb, @Nonnull CellStyle style, @Nullable CellStyleDescriptor descriptor) {
    if (descriptor == null) {
      return;
    }

    Font font = wb.createFont();
    style.setFont(font);
    font.setColor(descriptor.getFgColor().getIndex());
    if (descriptor.getFontSize() != null) {
      font.setFontHeightInPoints(descriptor.getFontSize());
    }
    font.setFontName(descriptor.getFontName());
    font.setItalic(descriptor.getFontStyles().contains(CellStyleDescriptor.FontStyle.ITALIC));
    if (descriptor.getFontStyles().contains(CellStyleDescriptor.FontStyle.BOLD)) {
      font.setBoldweight(CellCreator.BOLDWEIGHT_BOLD);
    }
  }

  public static void setBorders(@Nonnull CellStyle cellStyle, @Nullable Border border) {
    if (border == null) {
      return;
    }
    cellStyle.setBorderLeft(border.getLeft());
    cellStyle.setBorderRight(border.getRight());
    cellStyle.setBorderBottom(border.getDown());
    cellStyle.setBorderTop(border.getUp());
  }

  // region support
  private static void setBorders(@Nonnull MergedRegion region, @Nullable Border border) {
    if (border == null || (MergedRegion.isDebugEnabled() && region.getCounter() >= threshold)) {
      return;
    }

    setLeftBorder(region, border);
    setRightBorder(region, border);
    setBottomBorder(region, border);
    setTopBorder(region, border);
  }

  private static void setLeftBorder(@Nonnull MergedRegion region, @Nonnull Border border) {
    int rowStart = region.getRow().getRowNum() + 1; // +1 to not set style on root cell
    int rowEnd = region.getRowEnd();
    int column = region.getColStart();

    for (int i = rowStart; i <= rowEnd; i++) {
      Cell cell = getOrCreateCell(region.getRow().getSheet(), i, column);
      if (!isRootCell(region, cell)) {
        setCellStyle(cell, null, border, null);
      }
    }
  }

  private static void setTopBorder(@Nonnull MergedRegion region, @Nonnull Border border) {
    int colStart = region.getColStart();
    int colEnd = region.getColEnd();
    int rowIndex = region.getRow().getRowNum();

    for (int i = colStart; i <= colEnd; i++) {
      Cell cell = getOrCreateCell(region.getRow().getSheet(), rowIndex, i);
      if (!isRootCell(region, cell)) {
        setCellStyle(cell, null, border, null);
      }
    }
  }

  private static void setRightBorder(@Nonnull MergedRegion region, @Nonnull Border border) {
    int rowStart = region.getRow().getRowNum();
    int rowEnd = region.getRowEnd();
    int column = region.getColEnd();

    for (int i = rowStart; i <= rowEnd; i++) {
      Cell cell = getOrCreateCell(region.getRow().getSheet(), i, column);
      if (!isRootCell(region, cell)) {
        setCellStyle(cell, null, border, null);
      }
    }
  }

  private static void setBottomBorder(@Nonnull MergedRegion region, @Nonnull Border border) {
    int colStart = region.getColStart();
    int colEnd = region.getColEnd();
    int rowIndex = region.getRowEnd();

    for (int i = colStart; i <= colEnd; i++) {
      Cell cell = getOrCreateCell(region.getRow().getSheet(), rowIndex, i);
      if (!isRootCell(region, cell)) {
        setCellStyle(cell, null, border, null);
      }
    }
  }

  /**
   * Tests whether the specified cell is the root cell of the specified MergeRegion
   */
  private static boolean isRootCell(@Nonnull MergedRegion region, @Nonnull Cell cell) {
    return cell.getRow().getRowNum() == region.getRow().getRowNum()
      && cell.getColumnIndex() == region.getColStart();
  }

  private static Cell getOrCreateCell(Sheet sheet, int rowIndex, int column) {
    return getOrCreateCell(getOrCreateRow(sheet, rowIndex), column);
  }

  private static Row getOrCreateRow(Sheet sheet, int rowIndex) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      return sheet.createRow(rowIndex);
    }
    return row;
  }

  private static Cell getOrCreateCell(Row row, int column) {
    Cell cell = row.getCell(column);
    if (cell == null) {
      return row.createCell(column);
    }
    return cell;
  }

  private static void setBGColor(@Nonnull CellStyle style, @Nullable CellStyleDescriptor descriptor) {
    if (descriptor != null && descriptor.getBgColor() != null) {
      style.setFillPattern(CellStyle.SOLID_FOREGROUND);
      style.setFillForegroundColor(descriptor.getBgColor().getIndex());
    }
  }

  /**
   * Warning: merged region must be created <strong>after</strong> borders have been set with method
   * {@link #setBorders(MergedRegion, Border)}.
   */
  private static void createMergedRegion(MergedRegion region) {
    if (MergedRegion.isDebugEnabled() && region.getCounter() >= threshold) {
      return;
    }
    LOGGER.debug("creating region {} for threshold {}", region.getCounter(), threshold);
    CellRangeAddress range = new CellRangeAddress(region.getRow().getRowNum(), region.getRowEnd(), region.getColStart(), region.getColEnd());
    region.getRow().getSheet().addMergedRegion(range);
  }

  // RichTextString
  public static Cell createCell(Row row, int columnIndex, RichTextString value,
                                @Nullable CellStyleDescriptor style, @Nullable Border border, @Nullable DataFormat dataformat) {
    Cell cell = row.createCell(columnIndex);
    cell.setCellValue(value);
    setCellStyle(cell, style, border, dataformat);
    return cell;
  }

  public static Cell createCell(Row row, int columnIndex, RichTextString value, CellStyleDescriptor style, Border border) {
    return createCell(row, columnIndex, value, style, border, null);
  }

  public static Cell createCell(Row row, int columnIndex, RichTextString value, CellStyleDescriptor style, DataFormat dataformat) {
    return createCell(row, columnIndex, value, style, null, dataformat);
  }

  public static Cell createCell(Row row, int columnIndex, RichTextString value, CellStyleDescriptor style) {
    return createCell(row, columnIndex, value, style, null, null);
  }

  public static Cell createCell(Row row, int columnIndex, RichTextString value) {
    return createCell(row, columnIndex, value, null, null, null);
  }

  /**
   * Creates a cell with a {@link RichTextString} value which spans over several merged cells (eg. a merged region).
   */
  public static Cell createCell(@Nonnull MergedRegion region, RichTextString value,
                                @Nullable CellStyleDescriptor style, @Nullable Border border, @Nullable DataFormat dataformat) {
    Cell cell = createCell(region.getRow(), region.getColStart(), value, style, border, dataformat);
    setBorders(region, border);
    createMergedRegion(region);
    return cell;
  }

  public static Cell createCell(@Nonnull MergedRegion region, RichTextString value, @Nullable CellStyleDescriptor style, @Nullable Border border) {
    return createCell(region, value, style, border, null);
  }

  public static Cell createCell(@Nonnull MergedRegion region, RichTextString value, @Nullable CellStyleDescriptor style, @Nullable DataFormat dataformat) {
    return createCell(region, value, style, null, dataformat);
  }

  public static Cell createCell(@Nonnull MergedRegion region, RichTextString value, @Nullable CellStyleDescriptor style) {
    return createCell(region, value, style, null, null);
  }

  public static Cell createCell(@Nonnull MergedRegion region, RichTextString value) {
    return createCell(region, value, null, null, null);
  }

  // String value
  public static Cell createCell(Row row, int columnIndex, String value, CellStyleDescriptor style, Border border, DataFormat dataformat) {
    Cell cell = row.createCell(columnIndex);
    cell.setCellValue(value);
    setCellStyle(cell, style, border, dataformat);
    return cell;
  }

  public static Cell createCell(Row row, int columnIndex, String value, CellStyleDescriptor style, Border border) {
    return createCell(row, columnIndex, value, style, border, null);
  }

  public static Cell createCell(Row row, int columnIndex, String value, CellStyleDescriptor style, DataFormat dataformat) {
    return createCell(row, columnIndex, value, style, null, dataformat);
  }

  public static Cell createCell(Row row, int columnIndex, String value, CellStyleDescriptor style) {
    return createCell(row, columnIndex, value, style, null, null);
  }

  public static Cell createCell(Row row, int columnIndex, String value) {
    return createCell(row, columnIndex, value, null, null, null);
  }

  /**
   * Creates a cell with a String value which spans over several merged cells (eg. a merged region).
   */
  public static Cell createCell(@Nonnull MergedRegion region, String value,
                                @Nullable CellStyleDescriptor style, @Nullable Border border, @Nullable DataFormat dataformat) {
    Cell cell = createCell(region.getRow(), region.getColStart(), value, style, border, dataformat);
    setBorders(region, border);
    createMergedRegion(region);
    return cell;
  }

  public static Cell createCell(@Nonnull MergedRegion region, String value, @Nullable CellStyleDescriptor style, @Nullable Border border) {
    return createCell(region, value, style, border, null);
  }

  public static Cell createCell(@Nonnull MergedRegion region, String value, @Nullable CellStyleDescriptor style, @Nullable DataFormat dataformat) {
    return createCell(region, value, style, null, dataformat);
  }

  public static Cell createCell(@Nonnull MergedRegion region, String value, @Nullable CellStyleDescriptor style) {
    return createCell(region, value, style, null, null);
  }

  public static Cell createCell(@Nonnull MergedRegion region, String value) {
    return createCell(region, value, null, null, null);
  }

  // float value
  public static Cell createCell(Row row, int columnIndex, float value, CellStyleDescriptor style, Border border, DataFormat dataformat) {
    Cell cell = row.createCell(columnIndex);
    cell.setCellValue(value);
    setCellStyle(cell, style, border, dataformat);
    return cell;
  }

  public static Cell createCell(Row row, int columnIndex, float value, CellStyleDescriptor style, Border border) {
    return createCell(row, columnIndex, value, style, border, null);
  }

  public static Cell createCell(Row row, int columnIndex, float value, CellStyleDescriptor style, DataFormat dataformat) {
    return createCell(row, columnIndex, value, style, null, dataformat);
  }

  public static Cell createCell(Row row, int columnIndex, float value, CellStyleDescriptor style) {
    return createCell(row, columnIndex, value, style, null, null);
  }

  public static Cell createCell(Row row, int columnIndex, float value) {
    return createCell(row, columnIndex, value, null, null, null);
  }

  /**
   * Creates a cell with a float value which spans over several merged cells (eg. a merged region).
   */
  public static Cell createCell(@Nonnull MergedRegion region, float value,
                                @Nullable CellStyleDescriptor style, @Nullable Border border, @Nullable DataFormat dataformat) {
    Cell cell = createCell(region.getRow(), region.getColStart(), value, style, border, dataformat);
    setBorders(region, border);
    createMergedRegion(region);
    return cell;
  }

  public static Cell createCell(@Nonnull MergedRegion region, float value, @Nullable CellStyleDescriptor style, @Nullable Border border) {
    return createCell(region, value, style, border, null);
  }

  public static Cell createCell(@Nonnull MergedRegion region, float value, @Nullable CellStyleDescriptor style, @Nullable DataFormat dataformat) {
    return createCell(region, value, style, null, dataformat);
  }

  public static Cell createCell(@Nonnull MergedRegion region, float value, @Nullable CellStyleDescriptor style) {
    return createCell(region, value, style, null, null);
  }

  public static Cell createCell(@Nonnull MergedRegion region, float value) {
    return createCell(region, value, null, null, null);
  }

  // int value
  public static Cell createCell(Row row, int columnIndex, int value, CellStyleDescriptor style, Border border, DataFormat dataformat) {
    Cell cell = row.createCell(columnIndex);
    cell.setCellValue(value);
    setCellStyle(cell, style, border, dataformat);
    return cell;
  }

  public static Cell createCell(Row row, int columnIndex, int value, CellStyleDescriptor style, Border border) {
    return createCell(row, columnIndex, value, style, border, null);
  }

  public static Cell createCell(Row row, int columnIndex, int value, CellStyleDescriptor style, DataFormat dataformat) {
    return createCell(row, columnIndex, value, style, null, dataformat);
  }

  public static Cell createCell(Row row, int columnIndex, int value, CellStyleDescriptor style) {
    return createCell(row, columnIndex, value, style, null, null);
  }

  public static Cell createCell(Row row, int columnIndex, int value) {
    return createCell(row, columnIndex, value, null, null, null);
  }

  /**
   * Creates a cell with a int value which spans over several merged cells (eg. a merged region).
   */
  public static Cell createCell(@Nonnull MergedRegion region, int value,
                                @Nullable CellStyleDescriptor style, @Nullable Border border, @Nullable DataFormat dataformat) {
    Cell cell = createCell(region.getRow(), region.getColStart(), value, style, border, dataformat);
    setBorders(region, border);
    createMergedRegion(region);
    return cell;
  }

  public static Cell createCell(@Nonnull MergedRegion region, int value, @Nullable CellStyleDescriptor style, @Nullable Border border) {
    return createCell(region, value, style, border, null);
  }

  public static Cell createCell(@Nonnull MergedRegion region, int value, @Nullable CellStyleDescriptor style, @Nullable DataFormat dataformat) {
    return createCell(region, value, style, null, dataformat);
  }

  public static Cell createCell(@Nonnull MergedRegion region, int value, @Nullable CellStyleDescriptor style) {
    return createCell(region, value, style, null, null);
  }

  public static Cell createCell(@Nonnull MergedRegion region, int value) {
    return createCell(region, value, null, null, null);
  }

  /**
   * Un pourcentage dans excel doit être exprimé en fraction (ie. 100% = 1, 5% = 0,05)
   */
  public static float formatRatio(float remise) {
    if (remise == 0f) {
      return remise;
    }
    if (remise >= 1) {
      return remise / 100f;
    }
    return remise;
  }
}
