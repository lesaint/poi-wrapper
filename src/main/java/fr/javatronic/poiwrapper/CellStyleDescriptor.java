package fr.javatronic.poiwrapper;

import com.google.common.collect.ImmutableSet;
import org.apache.poi.ss.usermodel.IndexedColors;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.util.Collections;
import java.util.Set;

import static fr.javatronic.poiwrapper.CellStyleDescriptor.FontStyle.ITALIC;
import static fr.javatronic.poiwrapper.CellStyleDescriptor.FontStyle.BOLD;
import static org.apache.poi.hssf.usermodel.HSSFFont.FONT_ARIAL;
import static org.apache.poi.ss.usermodel.CellStyle.VERTICAL_CENTER;
import static org.apache.poi.ss.usermodel.IndexedColors.BLACK;

/**
 * CellStyleDescriptor -
 *
 * @author Sébastien Lesaint
 */
public class CellStyleDescriptor {

  public static enum FontStyle {
    ITALIC, BOLD
  }

  private static final Set<CellStyleDescriptor.FontStyle> ITALIC_AND_BOLD_SET = ImmutableSet.of(ITALIC, BOLD);
  private static final Set<CellStyleDescriptor.FontStyle> ITALIC_SET = ImmutableSet.of(ITALIC);
  private static final Set<CellStyleDescriptor.FontStyle> BOLD_SET = ImmutableSet.of(BOLD);

  @Nonnull
  private final Set<FontStyle> fontStyles;
  @Nonnull
  private final String fontName;
  @Nullable
  private final Short fontSize;
  @Nullable
  private final Short alignH;
  private final short alignV;
  @Nullable
  private final Short rotation;
  @Nonnull
  private final IndexedColors fgColor;
  @Nullable
  private final IndexedColors bgColor;
  private final boolean wrapText;

  public static Builder builder() {
    return new Builder();
  }

  private CellStyleDescriptor(Builder builder) {
    this.fontStyles = buildFontStylesSet(builder.italic, builder.bold);
    this.fontName = defaultStringName(builder.fontName);
    this.fontSize = builder.fontSize;
    this.alignH = builder.alignH;
    this.alignV = defaultAlignV(builder.alignV);
    this.rotation = builder.rotation;
    this.fgColor = defaultFgColor(builder.fgColor);
    this.bgColor = builder.bgColor;
    this.wrapText = builder.wrapText;
  }

  private static Set<CellStyleDescriptor.FontStyle> buildFontStylesSet(boolean italic, boolean bold) {
    if (!italic && !bold) {
      return Collections.emptySet();
    }
    if (italic && bold) {
      return ITALIC_AND_BOLD_SET;
    }
    if (italic) {
      return ITALIC_SET;
    }
    return BOLD_SET;
  }

  @Nonnull
  private String defaultStringName(@Nullable String fontName) {
    return fontName == null ? FONT_ARIAL : fontName;
  }

  @Nonnull
  private short defaultAlignV(@Nullable Short alignV) {
    return alignV == null ? VERTICAL_CENTER : alignV;
  }

  @Nonnull
  private IndexedColors defaultFgColor(@Nullable IndexedColors fgColor) {
    return fgColor == null ? BLACK : fgColor;
  }

  @Nonnull
  public Set<FontStyle> getFontStyles() {
    return fontStyles;
  }

  @Nonnull
  public String getFontName() {
    return fontName;
  }

  @Nullable
  public Short getFontSize() {
    return fontSize;
  }

  @Nullable
  public Short getAlignH() {
    return alignH;
  }

  public short getAlignV() {
    return alignV;
  }

  @Nullable
  public Short getRotation() {
    return rotation;
  }

  @Nonnull
  public IndexedColors getFgColor() {
    return fgColor;
  }

  @Nullable
  public IndexedColors getBgColor() {
    return bgColor;
  }

  public boolean isWrapText() {
    return wrapText;
  }

  public static class Builder {
    private boolean italic = false;
    private boolean bold = false;
    @Nullable
    private String fontName;
    @Nullable
    private Short fontSize;
    @Nullable
    private Short alignH;
    private short alignV;
    @Nullable
    private Short rotation;
    @Nullable
    private IndexedColors fgColor;
    @Nullable
    private IndexedColors bgColor;
    private boolean wrapText;

    public Builder withItalic(boolean italic) {
      this.italic = italic;
      return this;
    }

    public Builder withBold(boolean bold) {
      this.bold = bold;
      return this;
    }

    public Builder withFontName(@Nullable String fontName) {
      this.fontName = fontName;
      return this;
    }

    public Builder withFontSize(@Nullable int fontSize) {
      this.fontSize = (short) fontSize;
      return this;
    }

    public Builder withAlignH(@Nullable Short alignH) {
      this.alignH = alignH;
      return this;
    }

    public Builder withAlignV(short alignV) {
      this.alignV = alignV;
      return this;
    }

    public Builder withRotation(@Nullable int rotation) {
      this.rotation = (short) rotation;
      return this;
    }

    public Builder withFgColor(@Nullable IndexedColors fgColor) {
      this.fgColor = fgColor;
      return this;
    }

    public Builder withBgColor(@Nullable IndexedColors bgColor) {
      this.bgColor = bgColor;
      return this;
    }

    public Builder withWrapText(boolean wrapText) {
      this.wrapText = wrapText;
      return this;
    }

    public CellStyleDescriptor build() {
      return new CellStyleDescriptor(this);
    }
  }
}
