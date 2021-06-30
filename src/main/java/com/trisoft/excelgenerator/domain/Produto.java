package com.trisoft.excelgenerator.domain;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class Produto {
	private long id;
	private String descricao;
	private double preco;
	private String lote;
	private LocalDate vencimento;
	private double medida;
	private UnidadeMedida unidadeMedida;
	private int quantidade;
	
	public Produto(long id, String descricao, Double preco, double medida,UnidadeMedida unidadeMedida, 
			int quantidade, String lote, LocalDate vencimento) {
		super();
		this.id = id;
		this.descricao = descricao;
		this.preco = preco;
		this.lote = lote;
		this.vencimento = vencimento;
		this.medida = medida;
		this.unidadeMedida = unidadeMedida;
		this.quantidade = quantidade;
	}

	public long getId() {
		return id;
	}

	public void setId(long id) {
		this.id = id;
	}

	public String getDescricao() {
		return descricao;
	}

	public void setDescricao(String descricao) {
		this.descricao = descricao;
	}

	public double getPreco() {
		return preco;
	}

	public void setPreco(double preco) {
		this.preco = preco;
	}

	public String getLote() {
		return lote;
	}

	public void setLote(String lote) {
		this.lote = lote;
	}

	public LocalDate getVencimento() {
		return vencimento;
	}

	public void setVencimento(LocalDate vencimento) {
		this.vencimento = vencimento;
	}

	public double getMedida() {
		return medida;
	}

	public void setMedida(double medida) {
		this.medida = medida;
	}

	public UnidadeMedida getUnidadeMedida() {
		return unidadeMedida;
	}

	public void setUnidadeMedida(UnidadeMedida unidadeMedida) {
		this.unidadeMedida = unidadeMedida;
	}

	public int getQuantidade() {
		return quantidade;
	}

	public void setQuantidade(int quantidade) {
		this.quantidade = quantidade;
	}
	
	public String toString() {
		return "\n[PRODUTO " + id + "]\nDescricao: " + descricao
				+ "\nPreço: R$" + preco
				+ "\nPeso: " + medida + " " + unidadeMedida
				+ "\nVencimento: " + vencimento.getDayOfMonth() + "/" + vencimento.getMonthValue() + "/" + vencimento.getYear();
				
	}
	
	public static List<Produto> getProdutos(){
		
		int ano = LocalDate.now().getYear();
		int mesAtual = LocalDate.now().getMonthValue();
		int proximoMes = LocalDate.now().getMonthValue() == 12 ? 1 : LocalDate.now().getMonthValue() + 1;
		int dia = LocalDate.now().getDayOfMonth();
		int proximos2Dias = LocalDate.now().getDayOfMonth() + 2 > 28 ? LocalDate.now().getDayOfMonth() : LocalDate.now().getDayOfMonth() + 2;
		
		Produto p001 = new Produto(1, "Arroz", 10.00, 5,UnidadeMedida.QUILO, 34, "0001", LocalDate.of(ano, mesAtual, proximos2Dias) );
		Produto p002 = new Produto(2, "Feijão", 7.00, 2, UnidadeMedida.QUILO, 25, "0001", LocalDate.of(ano, mesAtual, proximos2Dias));
		Produto p003 = new Produto(3, "Macarrão", 3.00, 500,UnidadeMedida.GRAMA, 13, "0001", LocalDate.of(ano, mesAtual, dia ) );
		Produto p004 = new Produto(4, "Sal", 8.00, 1,UnidadeMedida.QUILO, 20, "0001", LocalDate.of(ano, mesAtual, dia) );
		Produto p005 = new Produto(5, "Açucar", 8.00, 2,UnidadeMedida.QUILO, 20, "0001", LocalDate.of(ano, proximoMes, dia) );
		Produto p006 = new Produto(6, "Biscoito", 5.99, 120,UnidadeMedida.GRAMA, 100, "0001", LocalDate.of(ano, proximoMes, dia) );
		Produto p007 = new Produto(7, "Pão", 4.48, 450,UnidadeMedida.GRAMA,995, "0001", LocalDate.of(ano-1, proximoMes , dia) );
		Produto p008 = new Produto(8, "Mortadela", 3.75, 1,UnidadeMedida.GRAMA, 10, "0001", LocalDate.of(ano-1, proximoMes, dia));
		Produto p009 = new Produto(9, "Mussarela", 7.00, 1,UnidadeMedida.QUILO, 10, "0001", LocalDate.of(ano-1, proximoMes, dia) );
		Produto p010 = new Produto(10, "Queijo", 15.00, 1,UnidadeMedida.QUILO, 10, "0001", LocalDate.of(ano-1, proximoMes, dia) );
		Produto p011 = new Produto(11, "Chocolate", 4.95, 90,UnidadeMedida.GRAMA, 50, "0001", LocalDate.of(ano, proximoMes, dia) );
		Produto p012 = new Produto(12, "Heineken", 4.39, 330,UnidadeMedida.ML, 90, "0001", LocalDate.of(ano, proximoMes, dia) );
		Produto p013 = new Produto(13, "Suco", 8.99, 900,UnidadeMedida.ML, 20, "0001", LocalDate.of(ano, proximoMes, dia) );
		Produto p014 = new Produto(14, "Arroz", 10.20, 5,UnidadeMedida.QUILO, 40, "0002", LocalDate.of(ano, proximoMes, dia) );
		Produto p015 = new Produto(15, "Feijão", 7.98, 2,UnidadeMedida.QUILO, 44, "0002", LocalDate.of(ano, proximoMes, dia) );
		Produto p016 = new Produto(16, "Coca-cola", 12.00, 2,UnidadeMedida.LITRO, 30, "0001", LocalDate.of(ano,proximoMes, dia) );
		Produto p017 = new Produto(17, "Coca-cola", 7.00, 1.5,UnidadeMedida.LITRO, 25, "0001", LocalDate.of(ano, proximoMes, dia) );
				
		List<Produto> produtos = new ArrayList<Produto>();
		produtos.add(p001);
		produtos.add(p002);
		produtos.add(p003);
		produtos.add(p004);
		produtos.add(p005);
		produtos.add(p006);
		produtos.add(p007);
		produtos.add(p008);
		produtos.add(p009);
		produtos.add(p010);
		produtos.add(p011);
		produtos.add(p012);
		produtos.add(p013);
		produtos.add(p014);
		produtos.add(p015);
		produtos.add(p016);
		produtos.add(p017);
		
		return produtos;
	}
}