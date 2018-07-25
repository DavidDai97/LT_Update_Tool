import java.awt.*;

public class Item {
    private double price;
    private String buyer;
    private int averageLeadTime;
    private int processLeadTime;
    private String currency;
    private int purchaseTime;
    private String supplier;
    private String description;
    private int moq;
    private int totalQuantity;

    public Item(String newBuyer, int leadTime, String newCurrency, double newPrice, String newSupplier, int newProcessLeadTime, String newDescription, int newMOQ, int totalQuantity){
        this.buyer = newBuyer;
        this.averageLeadTime = leadTime;
        this.currency = newCurrency;
        this.purchaseTime = 1;
        this.price = newPrice;
        this.supplier = newSupplier;
        this.processLeadTime = newProcessLeadTime;
        this.description = newDescription;
        this.moq = newMOQ;
        this.totalQuantity = totalQuantity;
    }

    public double getPrice() {
        return price;
    }
    public int getAverageLeadTime() {
        return averageLeadTime;
    }
    public int getPurchaseTime() {
        return purchaseTime;
    }
    public String getBuyer() {
        return buyer;
    }
    public String getCurrency() {
        return currency;
    }
    public String getSupplier(){
        return this.supplier;
    }
    public int getProcessLeadTime(){
        return this.processLeadTime;
    }
    public String getDescription(){
        return this.description;
    }
    public int getColor(){
        if(processLeadTime > (averageLeadTime+20) || processLeadTime < (averageLeadTime-20)){
            return 1;
        }
        else if(processLeadTime > (averageLeadTime+10) || processLeadTime < (averageLeadTime-10)){
            return 2;
        }
        return 0;
    }
    public int getDifference(){
        return Math.abs(this.processLeadTime-this.averageLeadTime);
    }
    public int getMOQ(){
        return this.moq;
    }
    public int getTotalQuantity(){
        return this.totalQuantity;
    }

    public void setAverageLeadTime(int newDifference){
        this.averageLeadTime = (int) Math.ceil(((double)(this.averageLeadTime + newDifference)) / (double)purchaseTime);
    }
    public void setPrice(double newPrice){
        this.price = newPrice;
    }
    public void setBuyer(String newBuyer){
        this.buyer = newBuyer;
    }
    public void setSupplier(String newSupplier){
        this.supplier = newSupplier;
    }
    public void setCurrency(String newCurrency){
        this.currency = newCurrency;
    }
    public void addQuantity(int quantity){
        this.totalQuantity += quantity;
    }

    public void addPurchaseTime(){
        this.purchaseTime++;
    }

}
