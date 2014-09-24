package fr.javatronic.poiwrapper;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.io.IOException;
import java.io.InputStream;

import static com.google.common.base.Preconditions.checkNotNull;

/**
 * SheetWrapper -
 *
 * @author Sébastien Lesaint
 */
public class SheetWrapper {
    @Nonnull
    private final Sheet sheet;

    private SheetWrapper(@Nonnull Sheet sheet) {
        this.sheet = checkNotNull(sheet);
    }

    public static SheetWrapper wrap(@Nonnull Sheet sheet) {
        return new SheetWrapper(sheet);
    }

    public static SheetWrapper create(@Nonnull Workbook wb, @Nullable String sheetName, @Nullable ZoomFactor zoomFactor) {
        Sheet sheet = sheetName == null ? wb.createSheet() : wb.createSheet(sheetName);
        if (zoomFactor != null) {
            sheet.setZoom(zoomFactor.getNumerator(), zoomFactor.getDenominator());
        }
        return wrap(sheet);
    }

    @Nonnull
    public Sheet getSheet() {
        return sheet;
    }

    /**
     * A row must not be created twice, otherwise styles of already created cells may be lost (or at least their styles).
     */
    @Nonnull
    public Row getOrCreateRow(int rowNumber) {
        Row row = sheet.getRow(rowNumber);
        if (row == null) {
            return sheet.createRow(rowNumber);
        }
        return row;
    }

    public void setColumnWidth(int columnIndex, int excelSize) {
        sheet.setColumnWidth(columnIndex, (int) (excelSize * 256));
    }

    /**
     * Do not use this method if the created merged region needs borders, uses method
     * {@code createCell(MergedRegion region)} from class {@link CellCreator}.
     */
    public void createMergedRegion(MergedRegion region) {
        CellRangeAddress range = new CellRangeAddress(region.getRow().getRowNum(), region.getRowEnd(), region.getColStart(), region.getColEnd());
        region.getRow().getSheet().addMergedRegion(range);
    }

    public void insertLogo(InputStream inputStream, int col, int row, double resize) throws IOException {
        byte[] bytes = IOUtils.toByteArray(inputStream);
        Workbook workbook = sheet.getWorkbook();
        int picId = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        inputStream.close();

        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(col);
        anchor.setRow1(row);
        Picture logo = drawing.createPicture(anchor, picId);
        logo.resize(resize);
    }

}
