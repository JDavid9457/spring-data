package com.bolsadeideas.springboot.app.view.xlsx;

import com.bolsadeideas.springboot.app.models.entity.Factura;
import com.bolsadeideas.springboot.app.models.entity.ItemFactura;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.context.support.MessageSourceAccessor;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.LocaleResolver;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.Map;

@Component("factura/ver.xlsx")
public class FacturaXlsxView extends AbstractXlsxView {

    @Autowired
    private MessageSource messageSource;

    @Autowired
    private LocaleResolver localeResolver;

    @Override
    protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request,
                                      HttpServletResponse response) throws Exception {

        response.setHeader("Content-Disposition", "attachment; filename=\"factura_view.xlsx\"");
        Factura factura = (Factura) model.get("factura");
        Sheet sheet = workbook.createSheet("Factura de compra");

        MessageSourceAccessor mensajes =  getMessageSourceAccessor();


        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);

        cell.setCellValue(mensajes.getMessage("text.factura.ver.datos.cliente"));
        row = sheet.createRow(1);
        cell = row.createCell(0);
        cell.setCellValue(factura.getCliente().getNombre() +"" + factura.getCliente().getApellido());

        row = sheet.createRow(2);
        cell = row.createCell(0);
        cell.setCellValue(factura.getCliente().getEmail());

        sheet.createRow(4).createCell(0).setCellValue(mensajes.getMessage("text.factura.ver.datos.factura"));
        sheet.createRow(5).createCell(0).setCellValue(mensajes.getMessage("text.cliente.factura.folio") + ": " +factura.getId());
        sheet.createRow(6).createCell(0).setCellValue(mensajes.getMessage("text.cliente.factura.descripcion") + ": "+ factura.getDescripcion());
        sheet.createRow(7).createCell(0).setCellValue(mensajes.getMessage("text.cliente.factura.fecha") + ": "   + factura.getFecha());


        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.MEDIUM);
        style.setBorderTop(BorderStyle.MEDIUM);
        style.setBorderRight(BorderStyle.MEDIUM);
        style.setBorderLeft(BorderStyle.MEDIUM);
        style.setFillBackgroundColor(IndexedColors.GOLD.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle body = workbook.createCellStyle();
        body.setBorderBottom(BorderStyle.THIN);
        body.setBorderTop(BorderStyle.THIN);
        body.setBorderRight(BorderStyle.THIN);
        body.setBorderLeft(BorderStyle.THIN);


        Row header = sheet.createRow(9);
        header.createCell(0).setCellValue(mensajes.getMessage("text.factura.form.item.nombre"));
        header.createCell(1).setCellValue(mensajes.getMessage("text.factura.form.item.precio"));
        header.createCell(2).setCellValue(mensajes.getMessage("text.factura.form.item.cantidad"));
        header.createCell(3).setCellValue(mensajes.getMessage("text.factura.form.item.total"));

        header.getCell(0).setCellStyle(style);
        header.getCell(1).setCellStyle(style);
        header.getCell(2).setCellStyle(style);
        header.getCell(3).setCellStyle(style);

        int rownum = 10;
        for(ItemFactura itemFactura: factura.getItems()) {
            Row fila = sheet.createRow(rownum ++);
            cell = fila.createCell(0);
            cell.setCellValue(itemFactura.getProducto().getNombre());
            cell.setCellStyle(body);

            cell = fila.createCell(1);
            cell.setCellValue(itemFactura.getProducto().getPrecio());
            cell.setCellStyle(body);

            cell = fila.createCell(2);
            cell.setCellValue(itemFactura.getCantidad());
            cell.setCellStyle(body);

            cell = fila.createCell(3);
            cell.setCellValue(itemFactura.calcularImporte());
            cell.setCellStyle(body);
        }

        Row filatotal = sheet.createRow(rownum);
        cell =filatotal.createCell(2);
        cell.setCellValue(mensajes.getMessage("text.factura.form.total") + ": ");
        cell.setCellStyle(body);

        cell = filatotal.createCell(3);
        cell.setCellValue(factura.getTotal());
        cell.setCellStyle(body);
    }
}
