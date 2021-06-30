package com.trisoft.excelgenerator.resource;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import com.trisoft.excelgenerator.builder.ExcelBuilder;
import com.trisoft.excelgenerator.domain.Produto;

@RestController
@RequestMapping("/excel")
public class ExcelResource {

	ExcelBuilder excel = getExcel();

	@GetMapping("/lista-produtos")
	public ResponseEntity<StreamingResponseBody> excel() {
		// attachment no lugar de inline
		return ResponseEntity.ok().contentType(MediaType.APPLICATION_OCTET_STREAM)
				.header(HttpHeaders.CONTENT_DISPOSITION, "inline;filename=\"lista-produtos.xlsx\"")
				.body(excel.getWorkbook()::write);
	}

	private ExcelBuilder getExcel() {
		List<Produto> produtos = Produto.getProdutos();
		produtos = produtos.stream().filter(p -> p.getLote().equals("0001")).collect(Collectors.toList());

		try {
			ExcelBuilder excel = new ExcelBuilder();
			excel.createSheet("Planilha de produtos");
			
			excel.createRow().createCell(HorizontalAlignment.LEFT, true).setCellValue("TRISOFT");
			excel.createRow().createRow().createCell(HorizontalAlignment.LEFT, true)
				.setCellValue(String.format("PRODUTOS EM ESTOQUE - Produtos encontrados: %s", produtos.size()));
			excel.setMergedRegion(2, 2, 0, 6);

			excel.createRow().createCell(HorizontalAlignment.CENTER, true).setCellValue("Descrição")
					.createCell(HorizontalAlignment.CENTER, true).setCellValue("Lote")
					.createCell(HorizontalAlignment.CENTER, true).setCellValue("Vencimento")
					.createCell(HorizontalAlignment.CENTER, true).setCellValue("Medida")
					.createCell(HorizontalAlignment.CENTER, true).setCellValue("Unidade de medida")
					.createCell(HorizontalAlignment.CENTER, true).setCellValue("Quantidade")
					.createCell(HorizontalAlignment.CENTER, true).setCellValue("Preço");

			CellStyle style = excel.createStyle();
			for (Produto produto : produtos) {
				excel.createRow().createNoEmptyCell(style, produto.getDescricao(), HorizontalAlignment.LEFT)
						.createNoEmptyCell(style, produto.getLote(), HorizontalAlignment.LEFT)
						.createNoEmptyCell(style, produto.getVencimento(), HorizontalAlignment.LEFT)
						.createNoEmptyCell(style, produto.getMedida(), HorizontalAlignment.RIGHT)
						.createNoEmptyCell(style, produto.getUnidadeMedida(), HorizontalAlignment.LEFT)
						.createNoEmptyCell(style, produto.getQuantidade(), HorizontalAlignment.RIGHT)
						.createCurrencyCell(style, BigDecimal.valueOf(produto.getPreco()), true);
			}
			
			Double total = produtos.stream().mapToDouble(Produto::getPreco).sum();
			Integer quantidade = produtos.stream().mapToInt(Produto::getQuantidade).sum();

			excel.createRow().createCell(HorizontalAlignment.RIGHT, true).setCellValue("Totais")
				.setMergedRegion(excel.getCell().getRowIndex(), excel.getCell().getRowIndex(), 0, 4)
				.createCell().skipCells(3).createCell(HorizontalAlignment.RIGHT, true).setCellValue(Double.valueOf(quantidade))
				.createCellMonetarioBold(BigDecimal.valueOf(total));
			
			BigDecimal totalGeral = BigDecimal.valueOf(total).multiply(BigDecimal.valueOf(quantidade));
			excel.createRow().skipCells(5).createCell(HorizontalAlignment.RIGHT, true).createCellMonetarioBold(totalGeral);
							
			//excel.createRow().createRow().createCell(excel.createStyle())
			//		.setCellValue(String.format("Total de unidades filtradas: %s", produtos.size()));
			//excel.setMergedRegion(excel.getCell().getRowIndex(), excel.getCell().getRowIndex(), 0, 6);
			//excel.getWorkbook().close();
			
			excel.autosize(0, 8);
			return excel;
		} catch (Exception e) { //
			e.printStackTrace();
		}
		return null;
	}
}
