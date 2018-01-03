using System;

namespace Leebl.Objects
{
  public class Lead
  {
    public string Leadcategorie;
    public string Leadbron;
    public DateTime Date;

    public Bedrijfsgegevens Bedrijfsgegevens;
    public Contactpersoon Contactpersoon;
    public Projectgegevens Projectgegevens;
  }
}