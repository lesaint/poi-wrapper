package fr.javatronic.poiwrapper;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * DataFormats -
 *
 * @author Sébastien Lesaint
 */
public final class DataFormats {

    public static final DataFormat CURRENCY = new PreexistingDataFormat((short) 7);
    public static final DataFormat PERCENTAGE = new PreexistingDataFormat((short) 0xa);

    public static DataFormat customDataFormat(Workbook wb, String format) {
        return new CustomDataFormat(wb, format);
    }

    private static class PreexistingDataFormat implements DataFormat {
        private final short format;

        public PreexistingDataFormat(short format) {
            this.format = format;
        }

        @Override
        public short getFormat() {
            return format;
        }
    }

    private static class CustomDataFormat implements DataFormat {
        private final short format;

        public CustomDataFormat(Workbook wb, String format) {
            this.format = wb.createDataFormat().getFormat(format);
        }

        @Override
        public short getFormat() {
            return format;
        }
    }
}
