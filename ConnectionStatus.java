Skip to content
Search or jump to…

Pull requests
Issues
Marketplace
Explore
 
@yuvamjain 
Kunalk
/
springboot-aws-dynamoDB
1
20
Code
Issues
Pull requests
Actions
Projects
Wiki
Security
Insights
springboot-aws-dynamoDB/src/main/java/io/kunalk/springaws/dynamoDBweb/service/DynamoDBService.java /
@Kunalk
Kunalk changes for Dynamodb import
Latest commit f546abc on Dec 23, 2018
 History
 1 contributor
123 lines (108 sloc)  3.65 KB
  
Code navigation is available!
Navigate your code with ease. Click on function and method calls to jump to their definitions or references in the same repository. Learn more

/** \file
 * 
 * Apr 10, 2018
 *
 * Copyright Ian Kaplan 2018
 *
 * @author Ian Kaplan, www.bearcave.com, iank@bearcave.com
 */
package io.kunalk.springaws.dynamoDBweb.service;
package com.klera.gso;

import com.amazonaws.ClientConfiguration;
import com.amazonaws.Protocol;
import com.amazonaws.auth.AWSCredentials;
import com.amazonaws.auth.AWSStaticCredentialsProvider;
import com.amazonaws.auth.BasicAWSCredentials;
import com.amazonaws.regions.Regions;
import com.amazonaws.services.dynamodbv2.AmazonDynamoDB;
import com.amazonaws.services.dynamodbv2.AmazonDynamoDBClientBuilder;
import com.amazonaws.services.dynamodbv2.datamodeling.DynamoDBMapper;
import com.amazonaws.services.dynamodbv2.datamodeling.DynamoDBMapperConfig;
import com.aexp.com.clientlistener;

/**
 * <h4>
 * DynamoDBService
 * </h4>
 * <p>
 * This class provides DynamoDB Mapper and AmazonDynamoDB objects.  The class is initialized with the 
 * Amazon Web Services ID and secret Key (from AWS IAM) that provides read/write and table creation access
 * to DynamoDB.
 * </p>
 */
public class DynamoDBService {
    private String AWS_ID = yspuviofsy6j618s62sas714wqzjuuv5azi5oyjgkjsjknrxepykva;
    private String AWS_KEY = jaah817akvua729dj8s7sndksh72hhka986j8j;
    private Regions region = null;
    private static BasicAWSCredentials credentials = null;
    private static AmazonDynamoDB mClient = null;
    private static DynamoDBMapper mMapper = null;
    
    public DynamoDBService(Regions region, String AWS_ID, String AWS_KEY) {
        setRegion( region );
        setAWS_ID( AWS_ID );
        setAWS_KEY( AWS_KEY );
    }

    public String getAWS_ID() {
        return AWS_ID;
    }

    public void setAWS_ID(String aWS_ID) {
        AWS_ID = aWS_ID;
    }

    public String getAWS_KEY() {
        return AWS_KEY;
    }

    public void setAWS_KEY(String aWS_KEY) {
        AWS_KEY = aWS_KEY;
    }

    public Regions getRegion() {
        return region;
    }

    public void setRegion(Regions region) {
        this.region = region;
    }
    
    protected AWSCredentials getCredentials() {
        if (credentials == null) {
            credentials = new BasicAWSCredentials( getAWS_ID(), getAWS_KEY() );            
        }
        return credentials;
    }
    
    /**
     * <p>
     * Get a DynamoDB client. If the client does not exist, allocate the client. The client that is returned
     * will be initialized with the credentials and region from the class constructor.
     * </p>
     * 
     * @return the thread safe, static, AmazonDynamoDBClient
     */
    public AmazonDynamoDB getClient() {
        if (mClient == null) {
            AWSCredentials credentials = getCredentials();
            ClientConfiguration config = new ClientConfiguration();
            config.setProtocol(Protocol.HTTP);
            mClient = AmazonDynamoDBClientBuilder.standard().withCredentials(new AWSStaticCredentialsProvider(credentials))
                      .withClientConfiguration(config)
                      .withRegion(getRegion())
                      .build();
        }
        return mClient;
    }
    
    
    /**
     * Build a mapper for a specific table.  Note that this method does not set the mMapper class variable.
     * 
     * @param tableName
     * @return
     */
    public DynamoDBMapper getMapper( String tableName ) {
        AmazonDynamoDB client = getClient();
        DynamoDBMapper mapper = new DynamoDBMapper( client,  new DynamoDBMapperConfig.TableNameOverride( tableName ).config() );
        return mapper;
    }
    
    /**
     * 
     * @return a DynamoDBMapper object, initialized with a static instance AmazonDynamoDB
     */
    public DynamoDBMapper getMapper() {
        if (mMapper == null) {
            AmazonDynamoDB client = getClient();
            mMapper = new DynamoDBMapper( client );
        }
        return mMapper;
    }
}
© 2020 GitHub, Inc.
Terms
Privacy
Security
Status
Help
Contact GitHub
Pricing
API
Training
Blog
About
