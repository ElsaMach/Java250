package com.example.demo.repository;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface ArticleRepository extends JpaRepository<Article, Long> {

    @Query("SELECT a FROM Article a WHERE UPPER(a.libelle) LIKE UPPER(CONCAT('%',:query,'%'))")
    List<Article> findByQuery(String query);

}
