package com.exeleuploader.Repo;

import com.exeleuploader.Model.User;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface UserRepo extends JpaRepository <User, Long> {

    Optional<User> findByEmail(String email);

}
