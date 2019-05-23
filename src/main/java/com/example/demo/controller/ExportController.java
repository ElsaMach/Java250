package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.service.ClientService;
import com.example.demo.service.ExportService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * Controlleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;
    @Autowired
    private FactureService factureService;
    @Autowired
    private ExportService exportService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        exportService.clientsCSV(response.getWriter());
    }

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        List<Client> allClients = clientService.findAllClients();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellPrenom = headerRow.createCell(1);
        cellPrenom.setCellValue("Prénom");

        Cell cellNom = headerRow.createCell(2);
        cellNom.setCellValue("Nom");

        int iRow = 1;
        for (Client client : allClients) {
            Row row = sheet.createRow(iRow);

            Cell id = row.createCell(0);
            id.setCellValue(client.getId());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(client.getPrenom());

            Cell nom = row.createCell(2);
            nom.setCellValue(client.getNom());

            iRow = iRow + 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{id}/factures/xlsx")
    public void factureXLSXByClient(@PathVariable("id") Long clientId, HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures-client-" + clientId + ".xlsx\"");
        List<Facture> factures = factureService.findFacturesClient(clientId);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Facture");
        Row headerRow = sheet.createRow(0);

        Cell cellId = headerRow.createCell(0);
        cellId.setCellValue("Id");

        Cell cellTotal = headerRow.createCell(1);
        cellTotal.setCellValue("Prix Total");

        int iRow = 1;
        for (Facture facture : factures) {
            Row row = sheet.createRow(iRow);

            Cell id = row.createCell(0);
            id.setCellValue(facture.getId());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(facture.getTotal());

            iRow = iRow + 1;
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/factures/xlsx")
    public void facturesclientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");
        List<Client> allClients = clientService.findAllClients();

        //Création cdu workbook
        Workbook workbook = new XSSFWorkbook();

        //Boucle sur les clients
        for (Client client : allClients) {
            //Création de l'onglet Client avec toutes les info
            Sheet sheet = workbook.createSheet(client.getNom());
            Row headerRow = sheet.createRow(0);

            //Création des noms colonnes
            Cell cellNom = headerRow.createCell(0);
            cellNom.setCellValue("Nom");

            Cell cellPrenom = headerRow.createCell(1);
            cellPrenom.setCellValue("Prenom");

            Cell cellDateNaissance = headerRow.createCell(2);
            cellDateNaissance.setCellValue("DateNaissance");

            Cell cellId = headerRow.createCell(3);
            cellId.setCellValue("Id");

            Row row = sheet.createRow(1);

            //Récupération des info
            Cell nom = row.createCell(0);
            nom.setCellValue(client.getNom());

            Cell prenom = row.createCell(1);
            prenom.setCellValue(client.getPrenom());

            Cell dateNaissance = row.createCell(2);
            dateNaissance.setCellValue(client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")));

            Cell id = row.createCell(3);
            id.setCellValue(client.getId());

            //Récupération des facture par Id client
            List<Facture> factures = factureService.findFacturesClient(client.getId());
            //LocalDate now = LocalDate.now();

            //i nous permet d'avoir le nom du bon client
            int i =1;
            //Boucle sur les factures
            for (Facture facture : factures) {

                //Onglet facture
                Sheet sheet2 = workbook.createSheet("Facture de " + client.getNom() + i);
                Row headerRow2 = sheet2.createRow(1);

                //Création des noms des colonnes
                Cell cellProduit = headerRow2.createCell(0);
                cellProduit.setCellValue("Produit");

                Cell cellQuantite = headerRow2.createCell(1);
                cellQuantite.setCellValue("Quantité");

                Cell cellPrix = headerRow2.createCell(2);
                cellPrix.setCellValue("Prix");

                Cell cellSousTotal = headerRow2.createCell(3);
                cellSousTotal.setCellValue("Sous total");

                //Boucle  sur ligne facture
                int frow = 1;
                for (LigneFacture lignefacture : facture.getLigneFactures()) {
                    Row rowBis = sheet2.createRow(frow);

                    Cell cellArticle = rowBis.createCell(0);
                    cellArticle.setCellValue(lignefacture.getArticle().getLibelle());

                    Cell cellQuantiteArticle = rowBis.createCell(1);
                    cellQuantiteArticle.setCellValue(lignefacture.getQuantite());

                    Cell cellPrixArticle = rowBis.createCell(2);
                    cellPrixArticle.setCellValue(lignefacture.getArticle().getPrix());

                    Cell cellSousTotalArticle = rowBis.createCell(3);
                    cellSousTotalArticle.setCellValue(lignefacture.getSousTotal());

                    frow = frow + 1;
                }
                Row rowTotal = sheet2.createRow(frow);
                Cell cellPrixTotal = rowTotal.createCell(3);
                cellPrixTotal.setCellValue(facture.getTotal());

                //autosize des colonnes
                sheet2.autoSizeColumn(0);
                sheet2.autoSizeColumn(1);
                sheet2.autoSizeColumn(2);
                sheet2.autoSizeColumn(3);
                
                //Incrémentation de i
                i++ ;
            }

            //autosize des colonnes
            sheet.autoSizeColumn(0, true);
            sheet.autoSizeColumn(1, true);
            sheet.autoSizeColumn(2, true);
            sheet.autoSizeColumn(3, true);

        }

        workbook.write(response.getOutputStream());
        workbook.close();

    }
}
