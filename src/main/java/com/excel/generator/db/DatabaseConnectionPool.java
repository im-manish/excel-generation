package com.excel.generator.db;

import com.excel.generator.util.ExcelGeneratorConstant;
import com.zaxxer.hikari.HikariConfig;
import com.zaxxer.hikari.HikariDataSource;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.sql.Connection;
import java.sql.SQLException;
import java.util.Objects;
import java.util.Properties;

/**
 * @author 6118454 - [Manish Kumar]
 */
public class DatabaseConnectionPool {
    private static Logger logger = LogManager.getLogger();

    private static HikariConfig config = new HikariConfig();
    private static HikariDataSource hikariOnPremDataSource;
    private DatabaseConnectionPool() {
    }
    private static void initializeConnectionPool(Properties properties) {
        logger.debug("Entering OnPremConnectionPool createConnection method");
        try {
            config.setDriverClassName(properties.getProperty(ExcelGeneratorConstant.DRIVER));
            config.setJdbcUrl(properties.getProperty(ExcelGeneratorConstant.URL));
            config.setUsername(properties.getProperty(ExcelGeneratorConstant.USER_ID));
            config.setPassword(properties.getProperty(ExcelGeneratorConstant.PASSWORD));
            config.setMaximumPoolSize(Integer.parseInt(properties.getProperty(ExcelGeneratorConstant.MAX_POOL_SIZE, "10")));
            config.addDataSourceProperty(ExcelGeneratorConstant.CACHE_PREP_STMTS, "true");
            config.addDataSourceProperty(ExcelGeneratorConstant.PREP_STMT_CACHE_SIZE, "250");
            config.addDataSourceProperty(ExcelGeneratorConstant.PREP_STMT_CACHE_SQL_LIMIT, "2048");
            config.setPoolName(ExcelGeneratorConstant.CONNECTION_POOL);
            hikariOnPremDataSource = new HikariDataSource(config);
        } catch (Exception e) {
            logger.error("Error while initializing op prem db", e);
            throw new RuntimeException("Error while initializing op prem db", e);
        }
        logger.debug("Exiting OnPremConnectionPool createConnection method");
    }


    public synchronized static Connection getConnection(Properties properties) throws SQLException {
        if (Objects.isNull(hikariOnPremDataSource)) {
            logger.debug("Start creating connection pool");
            initializeConnectionPool(properties);
        }
        return hikariOnPremDataSource.getConnection();
    }


}
