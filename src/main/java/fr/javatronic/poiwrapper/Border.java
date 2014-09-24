package fr.javatronic.poiwrapper;

import org.apache.poi.ss.usermodel.CellStyle;

import javax.annotation.Nullable;

import static org.apache.poi.ss.usermodel.CellStyle.BORDER_MEDIUM;
import static org.apache.poi.ss.usermodel.CellStyle.BORDER_NONE;
import static org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN;

/**
* Border -
*
* @author Sébastien Lesaint
*/
public class Border {
    // Thin borders
    public static final Border ALL_THIN_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_THIN, BORDER_THIN);
    public static final Border DOWN_THIN_BORDER = new Border(BORDER_THIN, BORDER_NONE, BORDER_NONE, BORDER_NONE);
    public static final Border UP_THIN_BORDER = new Border(BORDER_NONE, BORDER_THIN, BORDER_NONE, BORDER_NONE);
    public static final Border RIGHT_THIN_BORDER = new Border(BORDER_NONE, BORDER_NONE, BORDER_THIN, BORDER_NONE);
    public static final Border LEFT_THIN_BORDER = new Border(BORDER_NONE, BORDER_NONE, BORDER_NONE, BORDER_THIN);

    public static final Border DOWN_UP_THIN_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_NONE, BORDER_NONE);
    public static final Border RIGHT_LEFT_THIN_BORDER = new Border(BORDER_NONE, BORDER_NONE, BORDER_THIN, BORDER_THIN);
    public static final Border DOWN_RIGHT_THIN_BORDER = new Border(BORDER_THIN, BORDER_NONE, BORDER_THIN, BORDER_NONE);

    public static final Border DOWN_UP_LEFT_THIN_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_NONE, BORDER_THIN);
    public static final Border DOWN_UP_RIGHT_THIN_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_THIN, BORDER_NONE);
    public static final Border DOWN_RIGHT_LEFT_THIN_BORDER = new Border(BORDER_THIN, BORDER_NONE, BORDER_THIN, BORDER_THIN);
    public static final Border UP_RIGHT_LEFT_THIN_BORDER = new Border(BORDER_NONE, BORDER_THIN, BORDER_THIN, BORDER_THIN);

    // Medium border
    public static final Border ALL_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM);
    public static final Border DOWN_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_NONE, BORDER_NONE, BORDER_NONE);
    public static final Border UP_MEDIUM_BORDER = new Border(BORDER_NONE, BORDER_MEDIUM, BORDER_NONE, BORDER_NONE);
    public static final Border RIGHT_MEDIUM_BORDER = new Border(BORDER_NONE, BORDER_NONE, BORDER_MEDIUM, BORDER_NONE);
    public static final Border LEFT_MEDIUM_BORDER = new Border(BORDER_NONE, BORDER_NONE, BORDER_NONE, BORDER_MEDIUM);

    public static final Border DOWN_LEFT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_NONE, BORDER_NONE, BORDER_MEDIUM);
    public static final Border DOWN_RIGHT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_NONE, BORDER_MEDIUM, BORDER_NONE);
    public static final Border UP_LEFT_MEDIUM_BORDER = new Border(BORDER_NONE, BORDER_NONE, BORDER_MEDIUM, BORDER_MEDIUM);

    public static final Border DOWN_UP_LEFT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_MEDIUM, BORDER_NONE, BORDER_MEDIUM);
    public static final Border DOWN_UP_RIGHT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM, BORDER_NONE);
    public static final Border DOWN_RIGHT_LEFT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_NONE, BORDER_MEDIUM, BORDER_MEDIUM);
    public static final Border UP_RIGHT_LEFT_MEDIUM_BORDER = new Border(BORDER_NONE, BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM);

    // Thin and Medium borders mixed
    public static final Border DOWN_THIN_LEFT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_NONE, BORDER_NONE, BORDER_MEDIUM);
    public static final Border DOWN_THIN_UP_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_MEDIUM, BORDER_NONE, BORDER_NONE);
    public static final Border DOWN_THIN_RIGHT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_NONE, BORDER_MEDIUM, BORDER_NONE);
    public static final Border DOWN_THIN_UP_LEFT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_MEDIUM, BORDER_NONE, BORDER_MEDIUM);
    public static final Border DOWN_THIN_UP_RIGHT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_MEDIUM, BORDER_MEDIUM, BORDER_NONE);
    public static final Border DOWN_THIN_UP_RIGHT_LEFT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM);

    public static final Border UP_THIN_DOWN_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_THIN, BORDER_NONE, BORDER_NONE);

    public static final Border UP_THIN_DOWN_LEFT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_THIN, BORDER_NONE, BORDER_MEDIUM);
    public static final Border UP_THIN_DOWN_RIGHT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_THIN, BORDER_MEDIUM, BORDER_NONE);
    public static final Border UP_THIN_DOWN_RIGHT_LEFT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_THIN, BORDER_MEDIUM, BORDER_MEDIUM);

    public static final Border DOWN_UP_THIN_LEFT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_NONE, BORDER_MEDIUM);
    public static final Border DOWN_UP_THIN_RIGHT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_MEDIUM, BORDER_NONE);

    public static final Border DOWN_UP_THIN_RIGHT_LEFT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_MEDIUM, BORDER_MEDIUM);

    public static final Border DOWN_UP_RIGHT_THIN_LEFT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_THIN, BORDER_MEDIUM);

    public static final Border DOWN_UP_LEFT_THIN_RIGHT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_THIN, BORDER_MEDIUM, BORDER_THIN);

    public static final Border DOWN_RIGHT_THIN_UP_LEFT_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_MEDIUM, BORDER_MEDIUM, BORDER_THIN);

    public static final Border DOWN_RIGHT_LEFT_THIN_UP_MEDIUM_BORDER = new Border(BORDER_THIN, BORDER_MEDIUM, BORDER_THIN, BORDER_THIN);

    public static final Border UP_RIGHT_THIN_DOWN_LEFT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_THIN, BORDER_THIN, BORDER_MEDIUM);

    public static final Border UP_RIGHT_LEFT_THIN_DOWN_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_THIN, BORDER_THIN, BORDER_THIN);

    public static final Border UP_LEFT_THIN_DOWN_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_THIN, BORDER_NONE, BORDER_THIN);

    public static final Border UP_LEFT_THIN_DOWN_RIGHT_MEDIUM_BORDER = new Border(BORDER_MEDIUM, BORDER_THIN, BORDER_MEDIUM, BORDER_THIN);

    private final short down;
    private final short up;
    private final short right;
    private final short left;

    private Border(@Nullable Short down, @Nullable Short up, @Nullable Short right, @Nullable Short left) {
        this.down = nonNullFrom(down);
        this.up = nonNullFrom(up);
        this.right = nonNullFrom(right);
        this.left = nonNullFrom(left);
    }

    private short nonNullFrom(@Nullable Short border) {
        return border == null ? CellStyle.BORDER_NONE : border;
    }

    public short getDown() {
        return down;
    }

    public short getUp() {
        return up;
    }

    public short getRight() {
        return right;
    }

    public short getLeft() {
        return left;
    }
}
