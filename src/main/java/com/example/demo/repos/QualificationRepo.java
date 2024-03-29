package com.example.demo.repos;

import com.example.demo.models.Qualification;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface QualificationRepo extends JpaRepository<Qualification, Integer> {
}
