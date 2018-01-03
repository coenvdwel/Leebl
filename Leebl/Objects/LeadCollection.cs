using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Leebl.Objects
{
  public class LeadCollection
  {
    public const char Separator = ';';
    private readonly List<Lead> _leads;

    public LeadCollection(string path)
    {
      _leads = new List<Lead>();

      if (!File.Exists(path)) File.WriteAllText(path, "");

      var skip = true;
      var lines = File.ReadAllLines(path);

      foreach (var line in lines)
      {
        if (skip)
        {
          skip = false;
          continue;
        }

        var s = line.Split(Separator);

        _leads.Add(new Lead
        {
          Leadcategorie = s[0],
          Leadbron = s[1],
          Date = DateTime.ParseExact(s[2], "yyyy-MM-dd", CultureInfo.InvariantCulture),
          Bedrijfsgegevens = new Bedrijfsgegevens
          {
            Naam = s[3],
            Adres = s[4],
            Postcode = s[5],
            Woonplaats = s[6],
            Postbus = s[7],
            PostbusPostcode = s[8],
            PostbusWoonplaats = s[9],
            Website = s[10],
            Oprichtingsjaar = s[11],
            Rechtsvorm = s[12],
            Branche = s[13],
            Kvk = s[14]
          },
          Contactpersoon = new Contactpersoon
          {
            Naam = s[15],
            Functie = s[16],
            Afdeling = s[17],
            Telefoon = s[18],
            Mobiel = s[19],
            Fax = s[20],
            Email = s[21],
          },
          Projectgegevens = new Projectgegevens
          {
            Bedrijfsactiviteiten = s[22],
            Branche = s[23],
            AantalMedewerkers = s[24],
            Website = s[25],
            Bijzonderheden = s[26],
            Rol = s[27],
            HuidigeSoftware = s[28],
            AantalGebruikersNieuweSoftware = s[29],
            AantalMedewerkersMagazijn = s[30],
            Magazijnoppervlakte = s[31],
            Beslissingstermijn = s[32],
            Informatie = s[33]
          }
        });
      }
    }

    public void Add(string message, DateTime date)
    {
      var leadcategorie = new Regex("Leadcategorie:(.*?)<").Match(message);
      var leadbron = new Regex("Leadbron:(.*?)<").Match(message);
      var naam = new Regex("Naam:(.*?)<").Match(message);
      var adres = new Regex("Adres:(.*?)<").Match(message);
      var postcode = new Regex("Postcode:(.*?)<").Match(message);
      var woonplaats = new Regex("Plaats:(.*?)<").Match(message);
      var postbus = new Regex("Postbus:(.*?)<").Match(message);
      var postbusPostcode = new Regex("Postbus postcode:(.*?)<").Match(message);
      var postbusPlaats = new Regex("Postbus plaats:(.*?)<").Match(message);
      var website = new Regex("Website:(.*?)<").Match(message);
      var oprichtingsjaar = new Regex("Oprichtingsjaar:(.*?)<").Match(message);
      var rechtsvorm = new Regex("Rechtsvorm:(.*?)<").Match(message);
      var branche = new Regex("Branche:(.*?)<").Match(message);
      var kvk = new Regex("KVK nr.:(.*?)<").Match(message);

      var naam2 = naam.NextMatch();
      var functie = new Regex("Functie:(.*?)<").Match(message);
      var afdeling = new Regex("Afdeling:(.*?)<").Match(message);
      var telefoon = new Regex("Telefoon:(.*?)<").Match(message);
      var mobiel = new Regex("Mobiele telefoon:(.*?)<").Match(message);
      var fax = new Regex("Faxnummer:(.*?)<").Match(message);
      var email = new Regex("E-mail:(.*?)<").Match(message);

      var bedrijfsactiviteiten = new Regex("Bedrijfsactiviteiten:(.*?)<").Match(message);
      var branche2 = branche.NextMatch();
      var aantalMedewerkers = new Regex("Aantal medewerkers:(.*?)<").Match(message);
      var website2 = website.NextMatch();
      var bijzonderheden = new Regex("Bijzonderheden over deze organisatie:(.*?)<").Match(message);
      var rol = new Regex("Uw rol in dit WMS project:(.*?)<").Match(message);
      var huidigeSoftware = new Regex("Huidige softwareoplossing:(.*?)<").Match(message);
      var aantalGebruikersNieuweSoftware = new Regex("Aantal gebruikers nieuwe oplossing:(.*?)<").Match(message);
      var aantalMedewerkersMagazijn = new Regex("Aantal medewerkers in magazijn:(.*?)<").Match(message);
      var magazijnoppervlakte = new Regex("Magazijnoppervlakte:(.*?)<").Match(message);
      var beslissingstermijn = new Regex("Beslissingstermijn:(.*?)<").Match(message);
      var informatie = new Regex("Specifieke informatie over dit project:(.*?)<").Match(message);

      #region Oud formaat

      if (!naam.Success)
      {
        var content = false;
        var lines = message.Split(new []{Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
        
        for (var i = 0; i < lines.Length; i++)
        {
          var line = lines[i].Trim();

          if (!content)
          {
            lines[i] = String.Empty;
            content = line.Contains("_________");
            continue;
          }

          if (line.Length < 1) continue;
          if (Char.IsUpper(line[0])) continue;

          lines[i - 1] += " " + line;
          lines[i] = String.Empty;
        }

        message = String.Join(Environment.NewLine, lines.Where(x => !String.IsNullOrEmpty(x)));

        leadcategorie = new Regex("^(.*?)$").Match("WMS leadgegevens (oud)");
        leadbron = new Regex("^(.*?)$").Match("Aanvrager WMS systemen (oud)");

        var aanhef = new Regex("Aanhef:(.*?)\r\n").Match(message);
        var voorletters = new Regex("Voorletters:(.*?)\r\n").Match(message);
        var achternaam = new Regex("Naam:(.*?)\r\n").Match(message);

        naam2 = new Regex("^(.*?)$").Match(String.Format("{0} {1} {2}",
          aanhef.Success ? aanhef.Groups[1].Value : String.Empty,
          voorletters.Success ? voorletters.Groups[1].Value : String.Empty,
          achternaam.Success ? achternaam.Groups[1].Value : String.Empty));

        functie = new Regex("Functie:(.*?)\r\n").Match(message);
        telefoon = new Regex("Telefoon bedrijf:(.*?)\r\n").Match(message);
        email = new Regex("Mailadres:(.*?)\r\n").Match(message);
        naam = new Regex("Bedrijfsnaam:(.*?)\r\n").Match(message);

        rol = new Regex("Welke rol heeft u in het WMS selectieproces\\?:(.*?)\r\n").Match(message);
        afdeling = new Regex("Afdeling:(.*?)\r\n").Match(message);
        postbus = new Regex("Postadres:(.*?)\r\n").Match(message);
        postbusPostcode = new Regex("Postadres postcode:(.*?)\r\n").Match(message);
        postbusPlaats = new Regex("Postadres plaats:(.*?)\r\n").Match(message);
        adres = new Regex("Vestigingsadres:(.*?)\r\n").Match(message);
        postcode = new Regex("Vestigingsadres postcode:(.*?)\r\n").Match(message);
        woonplaats = new Regex("Vestigingsadres plaats:(.*?)\r\n").Match(message);

        aantalMedewerkersMagazijn = new Regex("Hoeveel medewerkers werken er in het magazijn waarvoor het WMS bestemd is\\?:(.*?)\r\n").Match(message);
        aantalMedewerkers = new Regex("Hoeveel medewerkers heeft de organisatie \\(of de organisatie waarvoor u het WMS zoekt\\)\\?:(.*?)\r\n").Match(message);
        magazijnoppervlakte = new Regex("Hoe groot is de magazijnoppervlakte waarvoor het WMS bestemd is\\?:(.*?)\r\n").Match(message);
        bedrijfsactiviteiten = new Regex("Wat zijn de bedrijfsactiviteiten\\?:(.*?)\r\n").Match(message);
        website = new Regex("Website:(.*?)\r\n").Match(message);
        beslissingstermijn = new Regex("Wanneer verwacht u nieuwe WMS software te kiezen\\?:(.*?)\r\n").Match(message);
        informatie = new Regex("Aanvullende informatie:(.*?)\r\n").Match(message);
        bijzonderheden = new Regex("Voorwaarden:(.*?)$").Match(message);
      }

      #endregion

      var lead = new Lead
      {
        Leadcategorie = leadcategorie.Success ? leadcategorie.Groups[1].Value : String.Empty,
        Leadbron = leadbron.Success ? leadbron.Groups[1].Value : String.Empty,
        Date = date,
        Bedrijfsgegevens = new Bedrijfsgegevens
        {
          Naam = naam.Success ? naam.Groups[1].Value : String.Empty,
          Adres = adres.Success ? adres.Groups[1].Value : String.Empty,
          Postcode = postcode.Success ? postcode.Groups[1].Value : String.Empty,
          Woonplaats = woonplaats.Success ? woonplaats.Groups[1].Value : String.Empty,
          Postbus = postbus.Success ? postbus.Groups[1].Value : String.Empty,
          PostbusPostcode = postbusPostcode.Success ? postbusPostcode.Groups[1].Value : String.Empty,
          PostbusWoonplaats = postbusPlaats.Success ? postbusPlaats.Groups[1].Value : String.Empty,
          Website = website.Success ? website.Groups[1].Value : String.Empty,
          Oprichtingsjaar = oprichtingsjaar.Success ? oprichtingsjaar.Groups[1].Value : String.Empty,
          Rechtsvorm = rechtsvorm.Success ? rechtsvorm.Groups[1].Value : String.Empty,
          Branche = branche.Success ? branche.Groups[1].Value : String.Empty,
          Kvk = kvk.Success ? kvk.Groups[1].Value : String.Empty,
        },
        Contactpersoon = new Contactpersoon
        {
          Naam = naam2.Success ? naam2.Groups[1].Value : String.Empty,
          Functie = functie.Success ? functie.Groups[1].Value : String.Empty,
          Afdeling = afdeling.Success ? afdeling.Groups[1].Value : String.Empty,
          Telefoon = telefoon.Success ? telefoon.Groups[1].Value : String.Empty,
          Mobiel = mobiel.Success ? mobiel.Groups[1].Value : String.Empty,
          Fax = fax.Success ? fax.Groups[1].Value : String.Empty,
          Email = email.Success ? email.Groups[1].Value : String.Empty,
        },
        Projectgegevens = new Projectgegevens
        {
          Bedrijfsactiviteiten = bedrijfsactiviteiten.Success ? bedrijfsactiviteiten.Groups[1].Value : String.Empty,
          Branche = branche2.Success ? branche2.Groups[1].Value : String.Empty,
          AantalMedewerkers = aantalMedewerkers.Success ? aantalMedewerkers.Groups[1].Value : String.Empty,
          Website = website2.Success ? website2.Groups[1].Value : String.Empty,
          Bijzonderheden = bijzonderheden.Success ? bijzonderheden.Groups[1].Value : String.Empty,
          Rol = rol.Success ? rol.Groups[1].Value : String.Empty,
          HuidigeSoftware = huidigeSoftware.Success ? huidigeSoftware.Groups[1].Value : String.Empty,
          AantalGebruikersNieuweSoftware = aantalGebruikersNieuweSoftware.Success ? aantalGebruikersNieuweSoftware.Groups[1].Value : String.Empty,
          AantalMedewerkersMagazijn = aantalMedewerkersMagazijn.Success ? aantalMedewerkersMagazijn.Groups[1].Value : String.Empty,
          Magazijnoppervlakte = magazijnoppervlakte.Success ? magazijnoppervlakte.Groups[1].Value : String.Empty,
          Beslissingstermijn = beslissingstermijn.Success ? beslissingstermijn.Groups[1].Value : String.Empty,
          Informatie = informatie.Success ? informatie.Groups[1].Value : String.Empty
        }
      };

      lead.Bedrijfsgegevens.Naam = lead.Bedrijfsgegevens.Naam.Replace(';', ':');
      lead.Bedrijfsgegevens.Naam = lead.Bedrijfsgegevens.Naam.Trim();

      if (_leads.All(l => l.Bedrijfsgegevens.Naam != lead.Bedrijfsgegevens.Naam))
      {
        _leads.Add(lead);
      }
    }

    public void Save(string path)
    {
      var sb = new StringBuilder();

      var header = new[]
      {
        "Leadcategorie",
        "Leadbron",
        "Datum",
        "Bedrijf Naam",
        "Bedrijf Adres",
        "Bedrijf Postcode",
        "Bedrijf Woonplaats",
        "Bedrijf Postbus",
        "Bedrijf Postbus Postcode",
        "Bedrijf Postbus Woonplaats",
        "Bedrijf Website",
        "Bedrijf Oprichtingsjaar",
        "Bedrijf Rechtsvorm",
        "Bedrijf Branche",
        "Bedrijf Kvk",
        "Contact Naam",
        "Contact Functie",
        "Contact Afdeling",
        "Contact Telefoon",
        "Contact Mobiel",
        "Contact Fax",
        "Contact Email",
        "Bedrijfsactiviteiten",
        "Branche",
        "Aantal Medewerkers",
        "Website",
        "Bijzonderheden",
        "Rol",
        "Huidige Software",
        "Aantal Gebruikers NieuweSoftware",
        "Aantal Medewerkers Magazijn",
        "Magazijnoppervlakte",
        "Beslissingstermijn",
        "Informatie"
      };

      sb.AppendLine(String.Join(Separator.ToString(CultureInfo.InvariantCulture), header));

      foreach (var values in _leads.Select(l => new[]
      {
        l.Leadcategorie,
        l.Leadbron,
        l.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        l.Bedrijfsgegevens.Naam,
        l.Bedrijfsgegevens.Adres,
        l.Bedrijfsgegevens.Postcode,
        l.Bedrijfsgegevens.Woonplaats,
        l.Bedrijfsgegevens.Postbus,
        l.Bedrijfsgegevens.PostbusPostcode,
        l.Bedrijfsgegevens.PostbusWoonplaats,
        l.Bedrijfsgegevens.Website,
        l.Bedrijfsgegevens.Oprichtingsjaar,
        l.Bedrijfsgegevens.Rechtsvorm,
        l.Bedrijfsgegevens.Branche,
        l.Bedrijfsgegevens.Kvk,
        l.Contactpersoon.Naam,
        l.Contactpersoon.Functie,
        l.Contactpersoon.Afdeling,
        l.Contactpersoon.Telefoon,
        l.Contactpersoon.Mobiel,
        l.Contactpersoon.Fax,
        l.Contactpersoon.Email,
        l.Projectgegevens.Bedrijfsactiviteiten,
        l.Projectgegevens.Branche,
        l.Projectgegevens.AantalMedewerkers,
        l.Projectgegevens.Website,
        l.Projectgegevens.Bijzonderheden,
        l.Projectgegevens.Rol,
        l.Projectgegevens.HuidigeSoftware,
        l.Projectgegevens.AantalGebruikersNieuweSoftware,
        l.Projectgegevens.AantalMedewerkersMagazijn,
        l.Projectgegevens.Magazijnoppervlakte,
        l.Projectgegevens.Beslissingstermijn,
        l.Projectgegevens.Informatie
      }))
      {
        for (var i = 0; i < values.Length; i++)
        {
          values[i] = values[i].Replace(';', ':');
          values[i] = values[i].Trim();
        }

        sb.AppendLine(String.Join(Separator.ToString(CultureInfo.InvariantCulture), values));
      }

      File.WriteAllText(path, sb.ToString());
    }
  }
}