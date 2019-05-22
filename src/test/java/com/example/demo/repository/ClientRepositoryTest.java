package com.example.demo.repository;

import com.example.demo.entity.Client;
import org.assertj.core.api.Assertions;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.core.AutoConfigureCache;
import org.springframework.boot.test.autoconfigure.orm.jpa.DataJpaTest;
import org.springframework.boot.test.autoconfigure.orm.jpa.TestEntityManager;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.List;

import static org.junit.Assert.*;

@RunWith(SpringRunner.class)
@DataJpaTest
public class ClientRepositoryTest {

    public static final String ALEXANDRE = "Alexandre";

    @Autowired
    private ClientRepository clientRepository;

    @Autowired
    private TestEntityManager testEntityManager;

    @Test
    public void save() {
        //GIVEN
        Client client = new Client();
        client.setNom("PETRILLO");
        client.setPrenom(ALEXANDRE);

        //WHEN
        clientRepository.save(client);

        // THEN
        List<Client> allClients = clientRepository.findAll();
        Assertions.assertThat(allClients).hasSize(1);
        Client uniqueClient = allClients.get(0);
        Assertions.assertThat(uniqueClient.getId()).isNotNull();
        Assertions.assertThat(uniqueClient.getPrenom()).isEqualTo(ALEXANDRE);
        Assertions.assertThat(uniqueClient.getNom()).isEqualToIgnoringCase("peTRIllO");
        Assertions.assertThat(uniqueClient.getDateNaissance()).isNull();
    }
}