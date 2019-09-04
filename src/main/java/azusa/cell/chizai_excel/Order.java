package azusa.cell.chizai_excel;

public class Order {

	private String day;
    private String food;
    private int number;
    private String remarks;

    public Order(String day, String food, int number, String remarks) {
        super();
        this.day = day;
        this.food = food;
        this.number = number;
        this.remarks = remarks;
    }

    // コンソール上で読めるように
    @Override
    public String toString() {
        return "Order [day=" + day + ", food=" + food + ", number=" + number
                + ", remarks" + remarks + "]";
    }

    public String getDay() {
        return day;
    }

    public void setDay(String day) {
        this.day = day;
    }

    public String getFood() {
        return food;
    }

    public void setFood(String food) {
        this.food = food;
    }

    public int getNumber() {
        return number;
    }

    public void setNumber(int number) {
        this.number = number;
    }

    public String getRemarks() {
        return remarks;
    }

    public void setRemarks(String remarks) {
        this.remarks = remarks;
    }

}