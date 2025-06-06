package club.xsir.poi;

/**
 * @author rwx
 * @version 1.0
 * @description:
 * @date 2025/3/25 15:00
 */
public class DataStruct implements Cloneable{

    @FieldInfo(order = 1,name = "序号")
    private String num;

    @FieldInfo(order = 2,name = "分支机构名称")
    private String stypeName;

    @FieldInfo(order = 3,name = "入库仓库")
    private String inStock;

    @FieldInfo(order = 4,name = "客户编码")
    private String customer;

    @FieldInfo(order = 5,name = "成品仓")
    private String cpStock;

    @FieldInfo(order = 6,name = "供应商")
    private String supplier;

    @FieldInfo(order = 7,name = "备注1")
    private String remark1;

    @FieldInfo(order = 8,name = "备注2")
    private String remark2;

    public String getNum() {
        return num;
    }

    public void setNum(String num) {
        this.num = num;
    }

    public String getStypeName() {
        return stypeName;
    }

    public void setStypeName(String stypeName) {
        this.stypeName = stypeName;
    }

    public String getInStock() {
        return inStock;
    }

    public void setInStock(String inStock) {
        this.inStock = inStock;
    }

    public String getCustomer() {
        return customer;
    }

    public void setCustomer(String customer) {
        this.customer = customer;
    }

    public String getCpStock() {
        return cpStock;
    }

    public void setCpStock(String cpStock) {
        this.cpStock = cpStock;
    }

    public String getSupplier() {
        return supplier;
    }

    public void setSupplier(String supplier) {
        this.supplier = supplier;
    }

    public String getRemark1() {
        return remark1;
    }

    public void setRemark1(String remark1) {
        this.remark1 = remark1;
    }

    public String getRemark2() {
        return remark2;
    }

    public void setRemark2(String remark2) {
        this.remark2 = remark2;
    }

    @Override
    protected Object clone() throws CloneNotSupportedException {
        return super.clone();
    }

    @Override
    public String toString() {
        return "DataStruct{" +
                "num='" + num + '\'' +
                ", stypeName='" + stypeName + '\'' +
                ", inStock='" + inStock + '\'' +
                ", customer='" + customer + '\'' +
                ", cpStock='" + cpStock + '\'' +
                ", supplier='" + supplier + '\'' +
                ", remark1='" + remark1 + '\'' +
                ", remark2='" + remark2 + '\'' +
                '}';
    }
}
