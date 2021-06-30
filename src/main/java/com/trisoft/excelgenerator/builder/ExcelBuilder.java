package com.trisoft.excelgenerator.builder;

import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ExcelBuilder {
	public static final String FORMAT_PERCENTUAL = "#,##0.00%";
	public static final String FORMAT_MONETARIO = "#,##0.00";
	private static final int ROW_WINDOW = 5000; // Quantos rows ficam na memória
	public static final String FORMAT = "#,##0";

	private SXSSFWorkbook workbook;
	private SXSSFSheet sheet;
	private Integer sheetIndice;
	private Map<Integer, Integer> sheetAutoSizeRowRef = new HashMap<>(); // Referência para autosize da sheet (Linha de
																			// cabeçalho por exemplo)
	private SXSSFRow row;
	private SXSSFCell cell;
	private int rowIndice;
	private int cellIndice;
	private boolean autoSize;
	private Font font;

	public void setSheetAutoSizeRowRef(int sheetIndex, int sheetAutoSizeRowRef) {
		this.sheetAutoSizeRowRef.put(sheetIndex, sheetAutoSizeRowRef);
	}

	public Integer getSheetIndice() {
		return sheetIndice;
	}

	public SXSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	public ExcelBuilder() {
		this.workbook = new SXSSFWorkbook(ROW_WINDOW);
		this.rowIndice = 0;
		this.cellIndice = 0;
	}

	public ExcelBuilder adicionarLinha() {
		this.rowIndice++;
		return this;
	}

	public ExcelBuilder adicionarLinha(int linhas) {
		this.rowIndice += linhas;
		return this;
	}

	public ExcelBuilder adicionarIndiceCell() {
		this.cellIndice++;
		return this;
	}

	public ExcelBuilder createSheet(String titulo) {
		this.cellIndice = 0;
		this.rowIndice = 0;
		this.sheetIndice = sheetIndice != null ? sheetIndice++ : 0;
		this.sheet = workbook.createSheet(titulo);
		this.sheet.trackAllColumnsForAutoSizing();
		return this;
	}

	public ExcelBuilder createRow(int linhas) {
		this.rowIndice += linhas;
		return createRowIndice(this.rowIndice);
	}

	public ExcelBuilder createRow() {
		return createRowIndice(this.rowIndice++);
	}

	public ExcelBuilder createRowIndice(int indice) {
		this.row = sheet.createRow(indice);
		this.font = null;
		this.cellIndice = 0;
		return this;
	}

	public ExcelBuilder createCell(int celulas) {
		skipCells(celulas);
		return createCell();
	}

	public ExcelBuilder createCell() {
		return createCell(row);
	}

	public ExcelBuilder createCell(SXSSFRow row) {
		this.cell = row.createCell(this.cellIndice++);
		if (this.autoSize) {
			autosize(this.cell.getColumnIndex());
		}

		this.font = createDefaultFont();
		return this;
	}

	public ExcelBuilder createCell(CellStyle style) {
		return createCell(style, row);
	}

	public ExcelBuilder createCell(CellStyle style, SXSSFRow row) {
		return createCell(style, HorizontalAlignment.LEFT, false, row);
	}

	public ExcelBuilder createCell(CellStyle style, HorizontalAlignment alignment, Boolean isBold) {
		return createCell(style, alignment, isBold, this.row);
	}

	public ExcelBuilder createCell(CellStyle style, HorizontalAlignment alignment, Boolean isBold, SXSSFRow row) {
		this.cell = row.createCell(this.cellIndice++);
		updateStyle(style, isBold, alignment);
		this.cell.setCellStyle(style);

		if (this.autoSize) {
			autosize(cell.getColumnIndex());
		}
		return this;
	}

	public ExcelBuilder createCell(CellStyle cellStyle, HorizontalAlignment horizontalAlign,
			VerticalAlignment verticalAlign) {
		createCell(cellStyle);

		CellStyle style = createStyle();
		style.cloneStyleFrom(cellStyle);

		style.setAlignment(horizontalAlign);
		style.setVerticalAlignment(verticalAlign);
		getCell().setCellStyle(style);
		return this;
	}

	public ExcelBuilder createCell(CellStyle cellStyle, HorizontalAlignment horizontalAlign) {
		createCell(cellStyle);

		CellStyle style = createStyle();
		style.cloneStyleFrom(cellStyle);

		style.setAlignment(horizontalAlign);
		getCell().setCellStyle(style);
		return this;
	}

	public ExcelBuilder createCells(List<String> cellValues, CellStyle cellStyle, HorizontalAlignment horizontalAlign,
			VerticalAlignment verticalAlign) {
		for (String cellValue : cellValues) {
			createCell(cellStyle, horizontalAlign, verticalAlign).getCell().setCellValue(cellValue);
		}
		return this;
	}

	public ExcelBuilder createCells(List<String> cellValues, CellStyle cellStyle, HorizontalAlignment horizontalAlign) {
		for (String cellValue : cellValues) {
			createCell(cellStyle, horizontalAlign).setCellValue(cellValue);
		}
		return this;
	}

	public ExcelBuilder createCellsBold(List<String> cellValues, HorizontalAlignment horizontalAlign) {
		for (String cellValue : cellValues) {
			createCell(horizontalAlign, true, row).setCellValue(cellValue);
		}
		return this;
	}

	public SXSSFCell getCell() {
		if (this.font != null) {
			CellStyle style = createStyle();
			style.cloneStyleFrom(this.cell.getCellStyle());
			style.setFont(this.font);
			this.cell.setCellStyle(style);
		}
		return this.cell;
	}

	public ExcelBuilder createCellRow(SXSSFSheet sheet, int rowNum, BigDecimal valor, boolean negate) {
		SXSSFRow row = sheet.getRow(rowNum);
		createCellMonetario(row, valor, negate, false);
		return this;
	}

	public ExcelBuilder createCellRow(SXSSFSheet sheet, int rowNum, BigDecimal valor) {
		return createCellRow(sheet, rowNum, valor, false);
	}

	public SXSSFCell createCellMonetario(SXSSFRow row, BigDecimal valor, Boolean negate, Boolean isBold) {
		SXSSFCell cell = row.createCell(row.getLastCellNum());
		CellStyle cellStyle = criarStyleNumber(isBold);
		if (Objects.isNull(valor)) {
			cell.setCellValue("-");
		} else {
			if (negate) {
				valor = valor.negate();
			}
			cell.setCellValue(valor.doubleValue());
		}
		cell.setCellStyle(cellStyle);
		if (this.autoSize) {
			autosize(cell.getColumnIndex());
		}
		return cell;
	}

	public ExcelBuilder createCellMonetario(BigDecimal valor) {
		this.cell = createCellMonetario(this.row, valor);
		return this;
	}

	public SXSSFCell createCellMonetario(SXSSFRow row, BigDecimal valor) {
		return createCellMonetario(row, valor, false);
	}

	public ExcelBuilder createCellMonetarioBold(BigDecimal valor) {
		this.cell = createCellMonetario(this.row, valor, true);
		return this;
	}

	public SXSSFCell createCellMonetario(SXSSFRow row, BigDecimal valor, Boolean isBold) {
		this.cellIndice++;
		return createCellMonetario(row, valor, false, isBold);
	}

	public ExcelBuilder createCellMonetarioNegate(BigDecimal valor, Boolean negate) {
		this.cell = createCellMonetario(row, valor, negate, false);
		return this;
	}

	public CellStyle criarStyleNumber(Boolean isBold) {
		return criarStyleNumber(isBold, 2);
	}

	public CellStyle criarStyleNumber(Boolean isBold, int casasDecimais) {
		this.font = createDefaultFont();

		CellStyle style = createStyle(isBold, HorizontalAlignment.RIGHT);
		style.setDataFormat(getFormatMonetario(casasDecimais));
		return style;
	}

	public short getFormatMonetario() {
		return getFormatMonetario(2);
	}

	public short getFormatMonetario(int casasDecimais) {
		String format = FORMAT + ".";
		for (int i = 0; i < casasDecimais; i++) {
			format += "0";
		}
		return this.workbook.createDataFormat().getFormat(format);
	}

	public ExcelBuilder styleBold() {
		this.font = createDefaultFont();// Create font
		this.font.setBold(true);// Make font bold
		return this;
	}

	public CellStyle createStyle() {
		return createStyle(false, HorizontalAlignment.LEFT);
	}

	public CellStyle createStyle(Boolean isBold, HorizontalAlignment alignment) {
		CellStyle cellStyle = this.workbook.createCellStyle();
		Font font = createDefaultFont();
		font.setBold(isBold);
		cellStyle.setFont(font);
		cellStyle.setAlignment(alignment);
		return cellStyle;
	}

	public void updateStyle(CellStyle style, Boolean isBold, HorizontalAlignment alignment) {
		font = createDefaultFont();
		font.setBold(isBold);
		if (style instanceof XSSFCellStyle) {
			XSSFFont fontStyle = ((XSSFCellStyle) style).getFont();
			if (fontStyle != null) {
				font.setColor(fontStyle.getColor());
			}
		}
		style.setFont(font);
		style.setAlignment(alignment);
	}

	public ExcelBuilder autosize(int colIndex) {
		return autosize(colIndex, false);
	}

	public ExcelBuilder autosize(int colIndex, boolean useMergedCells) {
		this.sheet.autoSizeColumn(colIndex, useMergedCells);
		return this;
	}

	public ExcelBuilder autosize(int fromColIndex, int toColIndex) {
		for (int i = fromColIndex; i <= toColIndex; i++) {
			this.autosize(i);
		}
		return this;
	}

	public ExcelBuilder setMergedRegion(int firstRow, int lastRow, int firstColumn, int lastColumn) {
		this.sheet.addMergedRegion(new CellRangeAddress(firstRow, // first row (0-based)
				lastRow, // last row (0-based)
				firstColumn, // first column (0-based)
				lastColumn// last column (0-based)
		));
		return this;
	}

	public ExcelBuilder mergeWithNextCell() {
		int prevRow = this.rowIndice - 1;
		int prevCol = this.cellIndice - 1;
		setMergedRegion(prevRow, prevRow, prevCol, this.cellIndice++);
		return this;
	}

	private CellStyle getFormatPercentual() {
		short format = this.workbook.createDataFormat().getFormat(FORMAT_PERCENTUAL);
		CellStyle cellStyleFormat = this.workbook.createCellStyle();
		cellStyleFormat.setDataFormat(format);
		cellStyleFormat.setAlignment(HorizontalAlignment.RIGHT);
		return cellStyleFormat;
	}

	public ExcelBuilder createCurrencyCell(CellStyle cellStyle, BigDecimal valor, Boolean zeroAsNull) {
		createCell(cellStyle);

		CellStyle _cellStyle = createStyle();
		_cellStyle.cloneStyleFrom(cellStyle);
		_cellStyle.setDataFormat(getFormatMonetario());

		if (Objects.isNull(valor) && zeroAsNull == false) {
			cell.setCellValue("-");
			_cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		} else if (Objects.isNull(valor) || (zeroAsNull == true && valor.compareTo(new BigDecimal(0)) == 0)) {
			cell.setCellValue("-");
			_cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		} else {
			cell.setCellValue(valor.doubleValue());
			cell.setCellType(CellType.NUMERIC);
			_cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		}
		cell.setCellStyle(_cellStyle);

		return this;
	}

	public ExcelBuilder createPercentCell(CellStyle cellStyle, BigDecimal valor, Boolean zeroAsNull) {
		createCell(cellStyle);

		CellStyle _cellStyle = createStyle();
		_cellStyle.cloneStyleFrom(cellStyle);
		short format = this.workbook.createDataFormat().getFormat(FORMAT_PERCENTUAL);
		_cellStyle.setDataFormat(format);

		if (Objects.isNull(valor) && zeroAsNull == false) {
			cell.setCellValue("-");
			_cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		} else if (Objects.isNull(valor) || (zeroAsNull == true && valor.compareTo(new BigDecimal(0)) == 0)) {
			cell.setCellValue("-");
			_cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		} else {
			valor = valor.setScale(10, BigDecimal.ROUND_CEILING);
			cell.setCellValue(valor.doubleValue());
			_cellStyle.setAlignment(HorizontalAlignment.RIGHT);
		}
		cell.setCellStyle(_cellStyle);

		return this;
	}

	public ExcelBuilder autosize() {
		this.autoSize = true;
		return this;
	}

	public ExcelBuilder createNoEmptyCell(Object valor, HorizontalAlignment horizontalAlignment) {
		return createNoEmptyCell(this.workbook.createCellStyle(), valor, horizontalAlignment);
	}

	public ExcelBuilder createNoEmptyCell(CellStyle cellStyle, Object valor, HorizontalAlignment horizontalAlignment) {
		this.cell = createCell(cellStyle).getCell();

		CellStyle noEmptyStyle = createStyle();
		noEmptyStyle.cloneStyleFrom(cellStyle);

		if (Objects.isNull(valor)) {
			this.cell.setCellValue("-");
		} else {
			this.cell.setCellValue(String.valueOf(valor));
		}
		noEmptyStyle.setAlignment(horizontalAlignment);

		this.cell.setCellStyle(noEmptyStyle);
		return this;
	}

	public ExcelBuilder skipCells(Integer cellsCount) {
		this.cellIndice += cellsCount;
		return this;
	}

	public CellStyle createStyleBold() {
		CellStyle cellStyle = createStyle();
		Font font = createDefaultFont();// Create font
		font.setBold(true);// Make font bold
		cellStyle.setFont(font);
		return cellStyle;
	}

	public ExcelBuilder setCellValue(Double value) {
		getCell().setCellValue(value);
		return this;
	}

	public ExcelBuilder setCellValue(String value) {
		getCell().setCellValue(value);
		return this;
	}

	public ExcelBuilder setCellValue(String value, Boolean notEmpty) {
		if (Objects.nonNull(notEmpty) && notEmpty) {
			value = Optional.ofNullable(value).orElse("-");
		}
		return setCellValue(value);
	}

	public ExcelBuilder createCell(HorizontalAlignment alignment) {
		return createCell(alignment, false);
	}

	public ExcelBuilder createCell(HorizontalAlignment alignment, Boolean isBold) {
		return createCell(alignment, isBold, row);
	}

	public ExcelBuilder createCell(HorizontalAlignment alignment, Boolean isBold, SXSSFRow row) {
		return createCell(this.workbook.createCellStyle(), alignment, isBold, row);
	}

	public Font getFont() {
		return this.font;
	}

	/**
	 * Padrão de fonte tamanho 10
	 * 
	 * @return
	 */
	public Font createDefaultFont() {
		Font font = this.workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		return font;
	}

	public ExcelBuilder createCellPercentual(Double percentual) {

		createCell(getFormatPercentual(), HorizontalAlignment.RIGHT, false, this.row);
		this.cell.setCellValue(percentual);
		return this;
	}

	public void autoSizeColumns() {
		int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			Integer refAutoSize = sheetAutoSizeRowRef.get(i);
			if (refAutoSize != null) {
				Row row = sheet.getRow(refAutoSize);
				if (row != null) {
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						int columnIndex = cell.getColumnIndex();
						sheet.autoSizeColumn(columnIndex);
					}
				}
			}
		}
	}

	// TODO
//	    @Getter
//	    @Setter
//	    public class ReferenciasGraficoExcelVO {
//
//	        private int linhaInicial = 0;
//	        private int linhaFinal = 0;
//	        private int colunaInicial = 0;
//	        private int colunaFinal = 0;
//
//	    }

//	    public void gerarGraficoColuna(String titulo, int colInicial, int rowInicial, int larguraGrafico, int alturaGrafico, ReferenciasGraficoExcelBuilderVO refGraficoLabels, ReferenciasGraficoExcelBuilderVO refGraficoValores) {
//	        //Verificar se funciona forçar o XSSF ao invés do SXSSF
//	        XSSFSheet sheet = this.workbook.getXSSFWorkbook().getSheetAt(sheetIndice);
//	        XSSFDrawing drawing = sheet.createDrawingPatriarch();
//	        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, colInicial, rowInicial, colInicial+larguraGrafico, rowInicial+alturaGrafico);
//	        XSSFChart chart = drawing.createChart(anchor);
//
//	        chart.setTitleText(titulo);
//	        chart.setTitleOverlay(false);
//
//	        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
//	        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
//	        leftAxis.getOrAddMajorGridProperties();
//	        leftAxis.getOrAddMinorGridProperties();
//	        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
//	        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
//
//	        XDDFDataSource<String> competencias = XDDFDataSourcesFactory.fromStringCellRange(sheet,
//	            new CellRangeAddress(refGraficoLabels.getLinhaInicial(), refGraficoLabels.getLinhaFinal(), refGraficoLabels.getColunaInicial(), refGraficoLabels.getColunaFinal()));
//
//	        XDDFNumericalDataSource<Double> valoresReais = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
//	            new CellRangeAddress(refGraficoValores.getLinhaInicial(), refGraficoValores.getLinhaFinal(), refGraficoValores.getColunaInicial(), refGraficoValores.getColunaFinal()));
//
//	        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
//	        XDDFChartData.Series series1 = data.addSeries(competencias, valoresReais);
//	        series1.setTitle(titulo, null);
//	        data.setVaryColors(true);
//	        XDDFBarChartData bar = (XDDFBarChartData) data;
//	        bar.setBarDirection(BarDirection.COL);
//	        chart.plot(data);
//	    }

//	    public ExcelBuilder getDefaultHeader(String sheetName, ExcelBuilder excel, Map<String, Object> parametros, Facade facade) {
//	        String titulo = (String) parametros.get(RelatorioUtils.NOME_RELATORIO);
//	        Condominio condominio = facade.service.condominio.getCondominioContexto();
//	        String nomeAdministradora = facade.service.relatorio.getNomeAdministradora(condominio.getParametroCondominio(), false);
//
//	        excel.createSheet(sheetName);
//	        if (!StringUtils.isEmpty(nomeAdministradora)) {
//	            excel.createRow().createCell().setCellValue(nomeAdministradora);
//	        }
//	        String nomeCondominio = facade.service.condominio.getCondominioContextoo().getNomeCondominio();
//	        excel.createRow().createCell(excel.createStyle()).setCellValue(nomeCondominio);
//	        excel.createRow().createRow().createCell(HorizontalAlignment.LEFT, true).setCellValue(titulo);
//	        return excel;
//	    }
//
//	    public ExcelBuilder getDefaultFooter( ExcelBuilder excel, String filtroRodape, Facade facade) {
//	        excel.createRow().createRow().createCell(excel.createStyle()).setCellValue(filtroRodape);
//	        String rodape = facade.service.relatorio.getRodapeCondominioAdministradora(facade.service.condominio.getCondominioContexto());
//	        excel.createRow().createRow().createCell(excel.createStyle()).setCellValue(rodape);
//	        return excel;
//	    }
}