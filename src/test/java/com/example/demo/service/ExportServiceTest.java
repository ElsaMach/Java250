package com.example.demo.service;

import com.example.demo.entity.Client;
import org.assertj.core.api.Assertions;
import org.hibernate.service.spi.InjectService;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.junit.MockitoJUnitRunner;
import org.springframework.boot.test.mock.mockito.MockBean;

import java.io.IOException;
import java.io.StringWriter;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.Collections;

import static java.util.Collections.singletonList;
import static org.junit.Assert.*;
import static org.mockito.Mockito.times;

@RunWith(MockitoJUnitRunner.class)
public class ExportServiceTest {

    @Mock
    private ClientService mockClientService;

    @InjectMocks
    private ExportService exportService;

    @Test
    public void clientsCSV_avecZeroClient() throws IOException {
        StringWriter stringWriter = new StringWriter();
        exportService.clientsCSV(stringWriter);
        Assertions.assertThat(stringWriter.toString()).contains("Id;Nom;");//...
    }

    @Test
    public void clientsCSV_avecUnClient() throws IOException {
        Client client = new Client();
        client.setNom("PETRILLO");
        client.setPrenom("Alexandre");
        client.setDateNaissance(LocalDate.now());

        Mockito.when(mockClientService.findAllClients()).thenReturn(singletonList(client));

        StringWriter stringWriter = new StringWriter();
        exportService.clientsCSV(stringWriter);
        Assertions.assertThat(stringWriter.toString()).contains("PETRILLO;Alexandre");//...

        Mockito.verify(mockClientService, times(1)).findAllClients();
    }
}