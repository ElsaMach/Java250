package com.example.demo.service;

import com.example.demo.entity.Client;
import org.assertj.core.api.Assertions;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.junit.MockitoJUnitRunner;

import java.io.IOException;
import java.io.StringWriter;
import java.time.LocalDate;

import static java.util.Collections.singletonList;

public class ExportServiceSansMockitoTest {


    @Test
    public void clientsCSV_avecUnClient() throws IOException {
        ClientServiceBidonImpl clientServiceBidon = new ClientServiceBidonImpl();
        ExportService exportService = new ExportService(clientServiceBidon);

        StringWriter stringWriter = new StringWriter();
        exportService.clientsCSV(stringWriter);
        Assertions.assertThat(stringWriter.toString()).contains("PETRILLO;Alexandre");//...


    }
}