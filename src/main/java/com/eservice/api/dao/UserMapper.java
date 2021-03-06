package com.eservice.api.dao;

import com.eservice.api.core.Mapper;
import com.eservice.api.model.user.User;
import com.eservice.api.model.user.UserDetail;
import org.apache.ibatis.annotations.Param;

import java.util.List;

public interface UserMapper extends Mapper<User> {
    UserDetail getUserAllDetail(@Param("id") Integer id );

    UserDetail requestLogin(@Param("account")String account, @Param("password")String password);

    List<UserDetail> selectUsers( @Param("account")String account,
                                  @Param("name")String name,
                                  @Param("roleId")Integer roleId,
                                  @Param("groupId")Integer groupId,
                                  @Param("marketGroupName")String marketGroupName,
                                  @Param("valid")Integer valid);
    User selectByAccount(@Param("account") String account);
    List<UserDetail> selectAllInstallGroupByUserId(@Param("id") Integer id);
    List<User> getUsersHaveOptimizePermit(@Param("thatPermit")String thatPermit);
}