
public class Eleve {
	
	String login;
	String email;
	int semaineInscription;
	int dernierHWReussi;
	
	public Eleve(String login, String email, int semaineInscription, int dernierHWReussi) {
		super();
		this.login = login;
		this.email = email;
		this.semaineInscription = semaineInscription;
		this.dernierHWReussi = dernierHWReussi;
	}
	
	public void afficherDetailEleve(){
		System.out.println("login:"+login+" email:"+email+" semaine d'inscription:"+semaineInscription+" dernier HW réussi:"+dernierHWReussi);
	}
}
