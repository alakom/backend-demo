package com.example.demo.repos;

import com.example.demo.models.Normas;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface NormasRepo extends JpaRepository<Normas, Integer> {
}
