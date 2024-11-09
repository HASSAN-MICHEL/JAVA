
import java.util.ArrayList;
import java.util.List;

public class Banque2{
    private static final double TAUX_PRIVE = 0.01;
    private static final double TAUX_EPARGNE = 0.02;

    private List<Client> clients;

    public Banque() {
        this.clients = new ArrayList<>();
    }

    public void ajouterClient(String nom, String ville, double soldePrive, double soldeEpargne) {
        Client client = new Client(nom, ville, soldePrive, soldeEpargne);
        clients.add(client);
    }

    public void bouclerComptes() {
        for (Client client : clients) {
            client.bouclerComptePrive(TAUX_PRIVE);
            client.bouclerCompteEpargne(TAUX_EPARGNE);
        }
    }

    public void afficherClients() {
        for (Client client : clients) {
            client.afficher();
        }
    }

    public static class Client {
        private String nom;
        private String ville;
        private Compte comptePrive;
        private Compte compteEpargne;

        public Client(String nom, String ville, double soldePrive, double soldeEpargne) {
            this.nom = nom;
            this.ville = ville;
            this.comptePrive = new Compte(soldePrive);
            this.compteEpargne = new Compte(soldeEpargne);
        }

        public void bouclerComptePrive(double taux) {
            comptePrive.boucler(taux);
        }

        public void bouclerCompteEpargne(double taux) {
            compteEpargne.boucler(taux);
        }

        public void afficher() {
            System.out.println("Client " + nom + " de " + ville);
            System.out.println(" Compte prive: " + comptePrive.getSolde() + " francs");
            System.out.println(" Compte d'epargne: " + compteEpargne.getSolde() + " francs");
        }
    }

    public static class Compte {
        private double solde;

        public Compte(double solde) {
            this.solde = solde;
        }

        public void boucler(double taux) {
            solde += taux * solde;
        }

        public double getSolde() {
            return solde;
        }
    }

    public static void main(String[] args) {
        Banque banque = new Banque();
        banque.ajouterClient("Pedro", "Geneve", 1000.0, 2000.0);
        banque.ajouterClient("Alexandra", "Lausanne", 3000.0, 4000.0);

        banque.afficherClients();

        banque.bouclerComptes();

        banque.afficherClients();
    }
}
