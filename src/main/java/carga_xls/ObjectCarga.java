package carga_xls;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import carga_xls.annotations.CabecalhoCarga;
import carga_xls.exceptions.CargaNoVariavelException;

public interface ObjectCarga<T> {

	default void trocaValor(String nomeCampo, Object valor,T objeto) throws IllegalArgumentException, IllegalAccessException {
		Class<? extends Object> class1 = objeto.getClass();
		Field field = null;
		try {
			field = localizarField(nomeCampo, class1);
		} catch (Exception e) {
			return;
		}
		field.setAccessible(true);
		if (valor instanceof String) {
			if (field.getType().equals(LocalDate.class)) {
				String valorTexto = (String) valor;
				LocalDate valorFormatado = null;
				if (valorTexto.isBlank()) {
					valorFormatado = LocalDate.parse(valorTexto);
				}
				field.set(objeto, valorFormatado);
			} else {
				field.set(objeto, (String) valor);
			}
		} else if (valor instanceof LocalDate) {

			field.set(objeto, (LocalDate) valor);

		} else if (valor instanceof Date) {
			LocalDate valorFormatado = ((Date) valor).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
			field.set(objeto, valorFormatado);
		}
		field.setAccessible(false);
	}
	default Field localizarField(String nomeCampo, Class classe) {
		Field[] declaredFields = classe.getDeclaredFields();
		for (Field f : declaredFields) {
			CabecalhoCarga annotation = f.getAnnotation(CabecalhoCarga.class);
			if (annotation != null) {
				if (annotation.name().trim().equalsIgnoreCase(nomeCampo.trim())
						|| f.getName().trim().equalsIgnoreCase(nomeCampo.trim())) {
					return f;
				}
			} else if (f.getName().trim().equalsIgnoreCase(nomeCampo.trim())) {
				return f;
			}
		}
		throw new CargaNoVariavelException("Nenhum atributo encontrado");
	}
	void cargaRowPlanilha(Row rowCabecalho, Row row);
	
	default void cargaRowPlanilha(Row rowCabecalho, Row row,T objeto)
			throws  SecurityException, IllegalArgumentException, IllegalAccessException {
		for (int i = 0; i < rowCabecalho.getLastCellNum(); i++) {
			Cell cellCabecalho = rowCabecalho.getCell(i);
			Cell cellValor = row.getCell(i);
			trocaValor(cellCabecalho.getStringCellValue(), carregaValorTipo(cellValor),objeto);
		}

	}

	default Object carregaValorTipo(Cell cellValor) {
		switch (cellValor.getCellType()) {
		case STRING:
			return cellValor.getStringCellValue();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cellValor)) {
				return  cellValor.getDateCellValue();
			}
			return cellValor.getNumericCellValue();
		default:
			return cellValor.getStringCellValue();
		}
	}
	
	

}
