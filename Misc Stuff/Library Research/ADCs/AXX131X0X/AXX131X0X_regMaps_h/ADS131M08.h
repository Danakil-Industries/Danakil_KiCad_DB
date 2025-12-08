//Define the registers, subregisters, and the start and length of each subregister

//Address 0x00 (ID REGISTER)
#DEFINE ADS131M08_REG_ID 0x00
//Define the subregisters
    //Bits 15-12 unused
    //Bits 11-8 are for RO_CHANCNT
        #DEFINE ADS131M08_REG_RO_CHANCNT 0x00
        #DEFINE ADS131M08_IDX_RO_CHANCNT 11
        #DEFINE ADS131M08_LEN_RO_CHANCNT 4
    //Bits 7-0 are unused