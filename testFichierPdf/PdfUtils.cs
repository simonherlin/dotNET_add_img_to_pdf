using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html;
using iTextSharp.text.html.simpleparser;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
//using System.Drawing;


namespace ARTEMIS.classes
{
    public class PdfUtils
    {
        private MemoryStream memStream;

        public static void AddImageOnEachPage(string cheminFichierFrom, string cheminFichierTo, string cheminImage)
        {
            // open the reader
            PdfReader reader = new PdfReader(cheminFichierFrom);
            Rectangle size = reader.GetPageSizeWithRotation(1);
            Document document = new Document(size);

            // open the writer
            FileStream fs = new FileStream(cheminFichierTo, FileMode.Create, FileAccess.Write);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();

            // the pdf content
            PdfContentByte cb = writer.DirectContent;

            if (!File.Exists(cheminImage))
                return;

            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(cheminImage);
            image.SetAbsolutePosition(0 + 20 + 1, HAUTEUR_MAX_PAGE - 25 - 20 - image.ScaledHeight);

            for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber++)
            {
                // create the new page and add it to the pdf
                PdfImportedPage page = writer.GetImportedPage(reader, pageNumber);
                cb.AddTemplate(page, 0, 0);

                //ajout de l'image
                cb.AddImage(image);

                if (pageNumber < reader.NumberOfPages)
                    document.NewPage();
            }

            document.Close();
        }


        public PdfUtils(string chemin, int orientation = ORIENTATION_PORTRAIT, bool useMemoryStream = false)
        {
            //constructeur
            cheminFichier = chemin;

            //creation du document pdf
            if (orientation == ORIENTATION_PORTRAIT)
            {
                largeurMax = LARGEUR_MAX_PAGE;
                hauteurMax = HAUTEUR_MAX_PAGE;
            }
            else
            {
                largeurMax = HAUTEUR_MAX_PAGE;
                hauteurMax = LARGEUR_MAX_PAGE;
            }
            document = new Document(new iTextSharp.text.Rectangle(0.0F, 0.0F, largeurMax, hauteurMax));

            Stream stream;
            //creation du fichier pdf
            if (useMemoryStream)
            {
                memStream = new MemoryStream();
                stream = memStream;
            }
            else
            {
                stream = new FileStream(cheminFichier, FileMode.Create);
            }

            //association du document au fichier
            PdfWriter writer = PdfWriter.GetInstance(document, stream);

            //ouverture du fichier
            document.Open();

            //recuperation du contenu du pdf
            cb = writer.DirectContent;

            //gestion pagination
            total = cb.CreateTemplate(100, 100);
            total.BoundingBox = new iTextSharp.text.Rectangle(-20, -20, 100, 100);


            //calcul des dimensions utiles
            largeurUtile = largeurMax - margeGauche - margeDroit;
            hauteurUtile = hauteurMax - margeHaut - margeBas;
        }

        //HACK use references to clean this up
        public PdfUtils(string chemin, int orientation, float x, float y) : this(chemin, orientation)
        {
            cb.AddTemplate(total, x + margeGauche + 1, y + margeBas);
            Texte("1 / ", x, hauteurUtile - 12 - y, 12, PdfUtils.FONT_HELVETICA, PdfUtils.ALIGNEMENT_DROIT);
        }

        private const float LARGEUR_MAX_PAGE = 593;
        private const float HAUTEUR_MAX_PAGE = 841;

        public const int ORIENTATION_PORTRAIT = 1;
        public const int ORIENTATION_PAYSAGE = 2;

        public const int ALIGNEMENT_GAUCHE = PdfContentByte.ALIGN_LEFT;
        public const int ALIGNEMENT_CENTRE = PdfContentByte.ALIGN_CENTER;
        public const int ALIGNEMENT_DROIT = PdfContentByte.ALIGN_RIGHT;

        public float largeurMax; //largeur max de la page sans tenir compte des marges
        public float hauteurMax; //hauteur max de la page sans tenir compte des marges

        public float margeGauche = 20; //marge gauche
        public float margeDroit = 20; //marge droite
        public float margeHaut = 20; //marge haute
        public float margeBas = 20; //marge basse

        public readonly float largeurUtile;
        public readonly float hauteurUtile;

        public float largeurLigne = 1;

        protected PdfContentByte cb; //contenu du pdf
        protected PdfTemplate total; //nb total de pages
        public Document document; //document
        protected string cheminFichier; //chemin du fichier de sortie

        //polices
        public const string FONT_HELVETICA = BaseFont.HELVETICA;
        public const string FONT_HELVETICA_BOLD = BaseFont.HELVETICA_BOLD;
        public const string FONT_HELVETICA_ITALIC = "Helvetica-Italic";
        public const string FONT_COURIER = BaseFont.COURIER;
        public const string FONT_TIMES_NEW_ROMAN = BaseFont.TIMES_ROMAN;
        public const string FONT_TIMES_ITALIC = BaseFont.TIMES_ITALIC;
        public const string FONT_SYMBOL = BaseFont.SYMBOL;
        public const string FONT_DINGBATS = BaseFont.ZAPFDINGBATS;

        public void Save()
        {
            //enregistrement des modifications
            if (document.PageNumber == 0)
                Texte("", 0, 0);
            document.Close();
        }

        public void SaveAvecTotal()
        {
            ecrireTotalPage();
            this.Save();
        }

        public void setLargeurLigne(float largeur)
        {
            largeurLigne = largeur;
        }

        public void Image(string cheminImage, float x, float y, float largeurAjustement, float hauteurAjustement)
        {
            //ajout d'une image en position absolute en pixel

            if (!File.Exists(cheminImage))
                return;

            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(cheminImage);

            //ajustement de la taille si demandé
            if (largeurAjustement != 0 && hauteurAjustement != 0)
                image.ScaleToFit(largeurAjustement, hauteurAjustement);

            //attention le point 0,0 est en bas a gauche
            //donc on recalcule pour simuler le point 0,0 en haut à gauche
            image.SetAbsolutePosition(x + margeGauche + 1, hauteurMax - y - margeHaut - image.ScaledHeight);

            //ajout de l'image
            cb.AddImage(image);
        }

        public void Image(string cheminImage, float x, float y, float tailleAjustement = 0)
        {
            Image(cheminImage, x, y, tailleAjustement, tailleAjustement);
        }

        public void TexteP(string texte, float gauche, float haut, float fontSize = 12, string fontName = "Courier")
        {
            //ecriture de texte en position absolute en %

            //calcul des positions réélles sur la page
            float x = (largeurUtile * gauche) / 100;
            float y = (largeurUtile * haut) / 100;

            Texte(texte, x, y, fontSize, fontName);
        }

        public void Texte(string texte, float x, float y, float fontSize = 12, string fontName = "Courier", int alignement = PdfUtils.ALIGNEMENT_GAUCHE, BaseColor couleur = null)
        {
            if (texte == null)
                return;
            //ecriture de texte en position absolute en pixel

            if (couleur == null)
                couleur = BaseColor.BLACK;

            //determination de la police
            string font = "Courier";
            switch (fontName.ToUpper())
            {
                case "HELVETICA":
                    font = BaseFont.HELVETICA;
                    break;
                case "HELVETICA-BOLD":
                    font = BaseFont.HELVETICA_BOLD;
                    break;
                case "HELVETICA-ITALIC":
                    font = BaseFont.HELVETICA_OBLIQUE;
                    break;
                case "COURIER":
                    font = BaseFont.COURIER;
                    break;
                case "TIMES-NEW-ROMAN":
                    font = BaseFont.TIMES_ROMAN;
                    break;
                case "ZAPFDINGBATS":
                    font = BaseFont.ZAPFDINGBATS;
                    break;
            }

            //on passe en mode texte
            cb.BeginText();

            //couleur du texte
            cb.SetColorFill(couleur);

            //creation de la police
            BaseFont bf = BaseFont.CreateFont(font, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            //affectation de la police au document
            cb.SetFontAndSize(bf, fontSize);

            //attention le point 0,0 est en bas a gauche
            //donc on recalcule pour simuler le point 0,0 en haut à gauche

            //ecriture du texte
            cb.ShowTextAligned(alignement, texte, x + margeGauche + 1, hauteurMax - y - margeHaut - fontSize, 0);

            //fin du mode texte
            cb.EndText();

        }

        public void TexteVertical(string texte, float x, float y, float fontSize, string fontName, int alignement, BaseColor couleur)
        {
            //ecriture de texte en position absolute en pixel

            //determination de la police
            string font = "Courier";
            switch (fontName.ToUpper())
            {
                case "HELVETICA":
                    font = BaseFont.HELVETICA;
                    break;
                case "HELVETICA-BOLD":
                    font = BaseFont.HELVETICA_BOLD;
                    break;
                case "HELVETICA-ITALIC":
                    font = BaseFont.HELVETICA_OBLIQUE;
                    break;
                case "COURIER":
                    font = BaseFont.COURIER;
                    break;
                case "TIMES-NEW-ROMAN":
                    font = BaseFont.TIMES_ROMAN;
                    break;
            }

            //on passe en mode texte
            cb.BeginText();

            //couleur du texte
            cb.SetColorFill(couleur);

            //creation de la police
            BaseFont bf = BaseFont.CreateFont(font, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            //affectation de la police au document
            cb.SetFontAndSize(bf, fontSize);

            //attention le point 0,0 est en bas a gauche
            //donc on recalcule pour simuler le point 0,0 en haut à gauche

            //ecriture du texte
            cb.ShowTextAligned(alignement, texte, x + margeGauche + 1, hauteurMax - y - margeHaut - fontSize, 90);

            //fin du mode texte
            cb.EndText();

        }

        public void TexteColonne(string texte, float x, float y, float fontSize, string fontName, int alignement, BaseColor couleur)
        {
            //determination de la police
            string font = "Courier";
            switch (fontName.ToUpper())
            {
                case "HELVETICA":
                    font = BaseFont.HELVETICA;
                    break;
                case "HELVETICA-BOLD":
                    font = BaseFont.HELVETICA_BOLD;
                    break;
                case "HELVETICA-ITALIC":
                    font = BaseFont.HELVETICA_OBLIQUE;
                    break;
                case "COURIER":
                    font = BaseFont.COURIER;
                    break;
                case "TIMES-NEW-ROMAN":
                    font = BaseFont.TIMES_ROMAN;
                    break;
            }
            //couleur du texte
            cb.SetColorFill(couleur);

            //creation de la police
            BaseFont bf = BaseFont.CreateFont(font, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);

            //affectation de la police au document
            cb.SetFontAndSize(bf, fontSize);

            ColumnText ct = new ColumnText(cb);
            ct.Alignment = alignement;
            ct.AddText(new Phrase(texte));
            ct.SetSimpleColumn(x + margeGauche + 1, hauteurMax - y - margeHaut - fontSize, 100, 30);
            ct.Go();
            cb.Stroke();
        }

        public void Rectangle(float x, float y, float largeur, float hauteur)
        {
            //affichage d'un rectangle en position absolute en pixel
            Rectangle(x, y, largeur, hauteur, 0, 0, 0);
        }

        public void Rectangle(float x, float y, float largeur, float hauteur, string bordureRGB)
        {
            //affichage d'un rectangle en position absolute en pixel
            Rectangle(x, y, largeur, hauteur, bordureRGB, "");
        }

        public void Rectangle(float x, float y, float largeur, float hauteur, string bordureRGB, string fondRGB)
        {
            //affichage d'un rectangle en position absolute en pixel

            //bordure
            int bRouge = 0;
            int bVert = 0;
            int bBleu = 0;
            if (bordureRGB != "")
            {
                bRouge = int.Parse(bordureRGB.Substring(0, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
                bVert = int.Parse(bordureRGB.Substring(2, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
                bBleu = int.Parse(bordureRGB.Substring(4, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
            }

            //fond
            int fRouge = 255;
            int fVert = 255;
            int fBleu = 255;
            if (fondRGB != "")
            {
                if (fondRGB == "-1")
                {
                    fRouge = -1;
                    fVert = -1;
                    fBleu = -1;
                }
                else
                {
                    fRouge = int.Parse(fondRGB.Substring(0, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
                    fVert = int.Parse(fondRGB.Substring(2, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
                    fBleu = int.Parse(fondRGB.Substring(4, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
                }
            }

            Rectangle(x, y, largeur, hauteur, bRouge, bVert, bBleu, fRouge, fVert, fBleu);

        }

        public void Rectangle(float x, float y, float largeur, float hauteur, int bRouge, int bVert, int bBleu, int fRouge, int fVert, int fBleu)
        {

            //affichage d'un rectangle en position absolute en pixel

            //selection de la couleur des traits
            cb.SetRGBColorStroke(bRouge, bVert, bBleu);
            cb.SaveState();
            cb.SetLineWidth(largeurLigne);

            //selection de la couleur de fond
            if (fRouge != -1 && fVert != -1 && fBleu != -1)
            {
                //couleur de fond normal
                cb.SetRGBColorFill(fRouge, fVert, fBleu);
            }
            else
            {
                //pas de couleur de fond, on met le fond transparent
                PdfGState gs = new PdfGState();
                gs.FillOpacity = 100;
                gs.FillOpacity = 0;
                cb.SetGState(gs);
            }

            //attention le point 0,0 est en bas a gauche
            //donc on recalcule pour simuler le point 0,0 en haut à gauche

            //affichage du rectangle
            cb.Rectangle(x + margeGauche + 1, hauteurMax - y - hauteur - margeHaut, largeur, hauteur);

            //validation du dessin
            cb.ClosePathFillStroke();
            cb.Stroke();

            cb.ResetRGBColorFill();
            cb.ResetRGBColorStroke();
            cb.RestoreState();
        }

        public void Rectangle(float x, float y, float largeur, float hauteur, int rouge, int vert, int bleu)
        {
            //affichage d'un rectangle en position absolute en pixel
            Rectangle(x, y, largeur, hauteur, rouge, vert, bleu, 255, 255, 255);
        }

        //REFACT: test float handles int case
        public void RectangleP(int gauche, int droite, int haut, int bas, int rouge = 0, int vert = 0, int bleu = 0)
        {
            //affichage d'un rectangle en position absolute en %

            //calcul des positions réélles sur la page
            float x = (largeurUtile * gauche) / 100;
            float largeur = ((largeurUtile * droite) / 100) - x;

            float y = (hauteurUtile * haut) / 100;
            float hauteur = ((hauteurUtile * bas) / 100) - y;

            Rectangle(x, y, largeur, hauteur, rouge, vert, bleu);
        }

        public void RectangleP(float gauche, float droite, float haut, float bas, int rouge = 0, int vert = 0, int bleu = 0)
        {
            //affichage d'un rectangle en position absolute en %

            //calcul des positions réélles sur la page
            float x = (largeurUtile * gauche) / 100;
            float largeur = ((largeurUtile * droite) / 100) - x;

            float y = (hauteurUtile * haut) / 100;
            float hauteur = ((hauteurUtile * bas) / 100) - y;

            Rectangle(x, y, largeur, hauteur, rouge, vert, bleu);
        }

        public void Ligne(float x1, float y1, float x2, float y2, string bordureRGB)
        {
            //affichage d'un trait en position absolute en pixel

            int bRouge = int.Parse(bordureRGB.Substring(0, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
            int bVert = int.Parse(bordureRGB.Substring(2, 2), System.Globalization.NumberStyles.AllowHexSpecifier);
            int bBleu = int.Parse(bordureRGB.Substring(4, 2), System.Globalization.NumberStyles.AllowHexSpecifier);

            //selection de la couleur des traits
            cb.SetRGBColorStroke(bRouge, bVert, bBleu);
            cb.SaveState();

            //attention le point 0,0 est en bas a gauche
            //donc on recalcule pour simuler le point 0,0 en haut à gauche

            //affichage du trait
            cb.MoveTo(x1 + margeGauche + 1, hauteurMax - y1 - margeHaut);
            cb.LineTo(x2 + margeGauche + 1, hauteurMax - y2 - margeHaut);

            cb.Stroke();

            cb.ResetRGBColorStroke();
            cb.RestoreState();
        }

        public void nouvellePage(int orientation = ORIENTATION_PORTRAIT)
        {
            //TODO gerer l'orientation de la page
            document.NewPage();
            this.nombrePages++;
        }

        //REFACT Cleanup code and remove useless parameters
        public void nouvellePageAvecNumerotation(int orientation = ORIENTATION_PORTRAIT, float x = 0, float y = 0, string fontName = FONT_HELVETICA, int fontSize = 12, string formatNumerotation = null, bool pourcentage = false)
        {
            document.NewPage();
            cb.AddTemplate(total, x + margeGauche + 1, margeBas + y);
            if (pourcentage)
            {
                this.AjouterNumerotation(x, y, formatNumerotation ?? "{0} / {1}", FontFactory.GetFont(FONT_HELVETICA, fontSize));
                return;
            }
            Texte(cb.PdfWriter.PageNumber + " / ", x, hauteurUtile - 12 - y, fontSize, fontName, PdfUtils.ALIGNEMENT_DROIT);
        }

        private class NumerotationUtils
        {
            public int pageIndex, numeroPage, nombrePages;
            public string format;
            public float x, y;
            public Font font;
        }

        private List<NumerotationUtils> pagesANumeroter = new List<NumerotationUtils>();
        public int nombrePages;
        /// <summary>
        /// Enregistre les informations de numérotation pour la page en cours.
        /// La numérotation est créé en appelant SaveEtNumeroter.
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontName"></param>
        private void AjouterNumerotation(float x, float y, string formatNumerotation, Font font)
        {
            pagesANumeroter.Add(
                new NumerotationUtils
                {
                    pageIndex = cb.PdfWriter.PageNumber,
                    numeroPage = ++nombrePages,
                    format = formatNumerotation,
                    x = x,
                    y = y,
                    font = font
                });
        }

        public void TerminerNumerotation()
        {
            int index = pagesANumeroter.Count - nombrePages;
            foreach (var pan in pagesANumeroter.GetRange(index, nombrePages))
            {
                pan.nombrePages = nombrePages;
            }
            nombrePages = 0;
        }

        /// <summary>
        /// Numérote les pages enregistrées avec AjouterNumerotation et enregistre le fichier.
        /// </summary>
        public void SaveEtNumeroter()
        {
            Save();

            var content = memStream.ToArray();
            var reader = new PdfReader(content);
            var stamper = new PdfStamper(reader, new FileStream(cheminFichier, FileMode.OpenOrCreate));
            foreach (var pan in pagesANumeroter)
            {
                var canvas = stamper.GetOverContent(pan.pageIndex);
                var phrase = new Phrase(string.Format(pan.format, pan.numeroPage, pan.nombrePages), pan.font);
                ColumnText.ShowTextAligned(canvas, Element.ALIGN_LEFT, phrase, pan.x, pan.y, 0);
            }
            stamper.Close();
            reader.Close();
        }

        public void ecrireTotalPage()
        {
            total.BeginText();
            total.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED), 12);
            total.SetTextMatrix(0, 0);
            int pageNumber = cb.PdfWriter.PageNumber;
            total.ShowText(pageNumber.ToString());
            total.EndText();
        }

        public static int getTaillePolicePourQueTexteRentre(String texte, int largeurPt, String police, int taillePoliceMax, int taillePoliceMin)
        {
            bool ok = false;
            while (taillePoliceMax > taillePoliceMin && !ok)
            {
                if (getLargeurTexte(texte, police, taillePoliceMax) < largeurPt)
                {
                    ok = true;
                }
                else
                {
                    taillePoliceMax -= 1;
                }
            }
            return taillePoliceMax;
        }

        public static float getLargeurTexte(String texte, String police, int taillePolice)
        {
            return new Chunk(texte, new Font(iTextSharp.text.pdf.BaseFont.CreateFont(police, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED), taillePolice)).GetWidthPoint();
        }

        public static List<String> decouperIntelligemmentChaine(String chaine, float largeurPt, String police, int taillePolice)
        {
            List<String> liste = new List<string>();
            if (chaine == null)
                return liste;
            while (getLargeurTexte(chaine, police, taillePolice) > largeurPt)
            {
                String sousChaine = chaine.TrimStart();
                do
                {
                    int lastSpace = sousChaine.LastIndexOf(' ');

                    if (lastSpace == -1)
                    {
                        sousChaine = sousChaine.Substring(0, sousChaine.Length - 1);
                    }
                    else
                    {
                        sousChaine = sousChaine.Substring(0, lastSpace);
                    }
                } while (getLargeurTexte(sousChaine, police, taillePolice) > largeurPt);
                liste.Add(sousChaine);
                chaine = chaine.Substring(sousChaine.Length).TrimStart();
            }
            liste.Add(chaine);
            return liste;
        }

    }
}
