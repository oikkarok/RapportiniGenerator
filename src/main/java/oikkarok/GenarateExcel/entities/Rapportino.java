package oikkarok.GenarateExcel.entities;

import java.util.HashMap;
import java.util.Map;

import lombok.Data;
import lombok.Getter;

@Data
public class Rapportino {
	
	public static final String proprietario = "Simone Coraccio";
	
	@Getter
    private static final Map<Progetto, Responsabile> mappaResponsabili = new HashMap<>();

	private String data;
	private int durataStimata;
	private int durataEffettiva;
	private String descrizione;
	private Progetto nomeProgetto;
    private Responsabile responsabile;
    
    // Mappa per associare i responsabili ai progetti
    static {
        mappaResponsabili.put(Progetto.NAQ, Responsabile.Autieri);
        mappaResponsabili.put(Progetto.NRB, Responsabile.Costagli);
        mappaResponsabili.put(Progetto.GTM, Responsabile.Caramia);
        mappaResponsabili.put(Progetto.NIF, Responsabile.Autieri);
        
        mappaResponsabili.put(Progetto.NED, Responsabile.Masi);
        mappaResponsabili.put(Progetto.NEC, Responsabile.Masi);
        mappaResponsabili.put(Progetto.CPA, Responsabile.Autieri);
        mappaResponsabili.put(Progetto.NRA, Responsabile.Autieri);
        mappaResponsabili.put(Progetto.GFP, Responsabile.Autieri);
        mappaResponsabili.put(Progetto.CLG, Responsabile.DiGrande);
        mappaResponsabili.put(Progetto.HSM, Responsabile.DiGrande);
        
    }
		
}
