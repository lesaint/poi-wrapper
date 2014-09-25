poi-wrapper
===========

A small wrapping layer around POI which makes creating OOXML Excel files with POI more readeable.

## Goal

By introducing various object and utility classes, this project aims at :

1. separating describing styles/font/dataformat from actually creating them
2. provide a readable way of creating cells and merged region in code (class to a DSL)

This allows factorize, reuse styles/font/border/dataformat in code and drastically reduce code duplication.

These informations can even be ```static```, allowing the use of ```constants``` and ```static imports``` drastically improving code readibility.

 - **Build Status:** [![Build Status](https://travis-ci.org/lesaint/poi-wrapper.svg)](https://travis-ci.org/lesaint/poi-wrapper)

## Limitations

This project is available as is and does not provides an extensive support of all styles/font/border/dataformat properties.

It only contains features that I had to use in a project of mine.

That said, the project is open source, feel free to fork or contribute to improve it.

 - **Issues:** https://github.com/mycila/license-maven-plugin/issues

## Documentation

The main entry point of using this wrapper is the [CellCreator](https://github.com/lesaint/poi-wrapper/blob/master/src/main/java/fr/javatronic/poiwrapper/CellCreator.java) class and its ```createCell``` static methods.

With the right static imports, creating a merged region with a border, a specific font and a format can be as consise and readable as the
following example.

```java
createCell(region(row, K, 14, L), totalPrice, black11CenterWhite, ALL_MEDIUM_BORDER, CURRENCY);
```

> creates a region spanning from cell ```12-K``` to ```14-L```, with a medium border all around it, in ```Arial``` 11, containing
> the total price formatted with excel default currency format

Please find below the relevant code requiered to achieve the writting of the code line above. It may seem like a lot of code for a single cell,
but there isn't many other lines imports required to create dozens of other cells.

```java
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static fr.javatronic.poiwrapper.Border.ALL_MEDIUM_BORDER;
import static fr.javatronic.poiwrapper.CellCreator.createCell;
import static fr.javatronic.poiwrapper.DataFormats.CURRENCY;
import static fr.javatronic.poiwrapper.MergedRegionFactory.region;
import static fr.javatronic.poiwrapper.Column.K;
import static fr.javatronic.poiwrapper.Column.L;
import static fr.javatronic.poiwrapper.SheetWrapper.create;
import static org.apache.poi.hssf.usermodel.HSSFFont.FONT_ARIAL;
import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER;
import static org.apache.poi.ss.usermodel.IndexedColors.BLACK;

[...]
    CellStyleDescriptor black11CenterWhite = CellStyleDescriptor.builder()
      .withFontName(FONT_ARIAL)
      .withFontSize(11)
      .withAlignH(ALIGN_CENTER)
      .withFgColor(BLACK)
      .build();

    Workbook wb = new XSSFWorkbook();
    SheetWrapper sheet = create(wb, "Name of the sheet here");
    Row row = sheet.getOrCreateRow(12);
    float totalPrice = computeTotalPrice();
    createCell(region(row, K, 14, L), totalPrice, black11CenterWhite, ALL_MEDIUM_BORDER, CURRENCY);

[...]
```