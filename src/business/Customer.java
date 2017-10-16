package business;


public class Customer {

    private String company;
    private String name;
    private String email;

    /**
     *
     * @param name
     * @param email
     */
    public Customer(String name, String email) {
        this.name = name;
        this.email = email;
    }

    /**
     *
     * @param company
     * @param name
     * @param email
     */
    public Customer(String company, String name, String email) {
        this.company = company;
        this.name = name;
        this.email = email;
    }

    public String getCompany() {
        return this.company;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public String getName() {
        return this.name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getEmail() { return this.email; }

    public void setEmail(String email) {
        this.email = email;
    }
}
