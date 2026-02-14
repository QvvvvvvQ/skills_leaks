// Palettes.cs - Document Color Scheme Library
//
// Usage Guidelines:
// - Text.* colors: Low saturation, high readability - use for ALL text content
// - Decorative.Muted.*: Subtle accents for professional/formal documents
// - Decorative.Vibrant.*: Bold accents for marketing/promotional materials
// - UI.*: Structural elements (borders, table headers, dividers)
//
// IMPORTANT: Never apply Decorative.Vibrant colors to body text

namespace DocumentFoundry;

public static class ColorSchemes
{
    // ════════════════════════════════════════════════════════════════════════
    // Verdant - Deep forest greens, editorial aesthetic
    // ════════════════════════════════════════════════════════════════════════
    public static class Verdant
    {
        public static class Text
        {
            public const string Primary = "1A3328";
            public const string Secondary = "2C4A3B";
            public const string Body = "3D5E4D";
            public const string Muted = "6B8B7A";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "4A7D62";
                public const string Secondary = "5E9178";
                public const string Accent = "72A58E";
            }
            public static class Vibrant
            {
                public const string Primary = "2D8B5A";
                public const string Secondary = "3AA06D";
                public const string Accent = "47B580";
            }
        }
        public static class UI
        {
            public const string Border = "A8C9B8";
            public const string TableHeader = "E3F0E8";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Slate - Cool gray tones, minimalist tech
    // ════════════════════════════════════════════════════════════════════════
    public static class Slate
    {
        public static class Text
        {
            public const string Primary = "1C2127";
            public const string Secondary = "2E353D";
            public const string Body = "424B56";
            public const string Muted = "6D7785";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "5A6472";
                public const string Secondary = "707B8A";
                public const string Accent = "8691A1";
            }
            public static class Vibrant
            {
                public const string Primary = "3B82C4";
                public const string Secondary = "4A94D6";
                public const string Accent = "59A6E8";
            }
        }
        public static class UI
        {
            public const string Border = "C4CCD6";
            public const string TableHeader = "EEF1F4";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Azure - Professional blue, corporate identity
    // ════════════════════════════════════════════════════════════════════════
    public static class Azure
    {
        public static class Text
        {
            public const string Primary = "0D1E36";
            public const string Secondary = "1A3254";
            public const string Body = "2A4670";
            public const string Muted = "5673A0";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "4068A0";
                public const string Secondary = "5580B8";
                public const string Accent = "6A98D0";
            }
            public static class Vibrant
            {
                public const string Primary = "1E6FD9";
                public const string Secondary = "2E82EB";
                public const string Accent = "4094FC";
            }
        }
        public static class UI
        {
            public const string Border = "B8CCE4";
            public const string TableHeader = "E8F0F8";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Amber - Warm earth tones, organic feel
    // ════════════════════════════════════════════════════════════════════════
    public static class Amber
    {
        public static class Text
        {
            public const string Primary = "3D2510";
            public const string Secondary = "5C3A1C";
            public const string Body = "7A5030";
            public const string Muted = "A07850";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "B08040";
                public const string Secondary = "C49450";
                public const string Accent = "D8A860";
            }
            public static class Vibrant
            {
                public const string Primary = "D97520";
                public const string Secondary = "E68A30";
                public const string Accent = "F39F40";
            }
        }
        public static class UI
        {
            public const string Border = "E4D0B8";
            public const string TableHeader = "FAF4EC";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Plum - Rich purple, creative industries
    // ════════════════════════════════════════════════════════════════════════
    public static class Plum
    {
        public static class Text
        {
            public const string Primary = "281830";
            public const string Secondary = "3E2848";
            public const string Body = "543860";
            public const string Muted = "7A5888";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "6B4880";
                public const string Secondary = "805898";
                public const string Accent = "9568B0";
            }
            public static class Vibrant
            {
                public const string Primary = "8B40C0";
                public const string Secondary = "A050D8";
                public const string Accent = "B560F0";
            }
        }
        public static class UI
        {
            public const string Border = "D4C0E0";
            public const string TableHeader = "F4EEF8";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Coral - Warm pink/salmon, lifestyle brands
    // ════════════════════════════════════════════════════════════════════════
    public static class Coral
    {
        public static class Text
        {
            public const string Primary = "3A1A20";
            public const string Secondary = "582830";
            public const string Body = "763840";
            public const string Muted = "A05860";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "C06870";
                public const string Secondary = "D07880";
                public const string Accent = "E08890";
            }
            public static class Vibrant
            {
                public const string Primary = "E84858";
                public const string Secondary = "F05868";
                public const string Accent = "F86878";
            }
        }
        public static class UI
        {
            public const string Border = "F0C8CC";
            public const string TableHeader = "FDF2F3";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Teal - Medical/clinical, trust signaling
    // ════════════════════════════════════════════════════════════════════════
    public static class Teal
    {
        public static class Text
        {
            public const string Primary = "0C2A2D";
            public const string Secondary = "184045";
            public const string Body = "24565D";
            public const string Muted = "408088";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "308890";
                public const string Secondary = "40A0A8";
                public const string Accent = "50B8C0";
            }
            public static class Vibrant
            {
                public const string Primary = "10A0B0";
                public const string Secondary = "20B4C4";
                public const string Accent = "30C8D8";
            }
        }
        public static class UI
        {
            public const string Border = "A8D8DC";
            public const string TableHeader = "E8F6F7";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Charcoal - Sophisticated monochrome, luxury
    // ════════════════════════════════════════════════════════════════════════
    public static class Charcoal
    {
        public static class Text
        {
            public const string Primary = "141414";
            public const string Secondary = "242424";
            public const string Body = "363636";
            public const string Muted = "5A5A5A";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "484848";
                public const string Secondary = "606060";
                public const string Accent = "787878";
            }
            public static class Vibrant
            {
                public const string Primary = "C4A040";
                public const string Secondary = "D4B050";
                public const string Accent = "E4C060";
            }
        }
        public static class UI
        {
            public const string Border = "C8C8C8";
            public const string TableHeader = "F0F0F0";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Crimson - Authority, formal documents
    // ════════════════════════════════════════════════════════════════════════
    public static class Crimson
    {
        public static class Text
        {
            public const string Primary = "2D0A10";
            public const string Secondary = "4A1620";
            public const string Body = "662430";
            public const string Muted = "8A4048";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "A03040";
                public const string Secondary = "B84050";
                public const string Accent = "D05060";
            }
            public static class Vibrant
            {
                public const string Primary = "D01830";
                public const string Secondary = "E02840";
                public const string Accent = "F03850";
            }
        }
        public static class UI
        {
            public const string Border = "E8C0C4";
            public const string TableHeader = "FCF0F1";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Olive - Sustainability, natural products
    // ════════════════════════════════════════════════════════════════════════
    public static class Olive
    {
        public static class Text
        {
            public const string Primary = "252818";
            public const string Secondary = "3A3E28";
            public const string Body = "505438";
            public const string Muted = "747858";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "6A7048";
                public const string Secondary = "808658";
                public const string Accent = "969C68";
            }
            public static class Vibrant
            {
                public const string Primary = "7A8830";
                public const string Secondary = "8CA040";
                public const string Accent = "9EB850";
            }
        }
        public static class UI
        {
            public const string Border = "CCD0B8";
            public const string TableHeader = "F2F4EC";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Indigo - Innovation, technology
    // ════════════════════════════════════════════════════════════════════════
    public static class Indigo
    {
        public static class Text
        {
            public const string Primary = "181830";
            public const string Secondary = "282848";
            public const string Body = "383860";
            public const string Muted = "585888";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "5050A0";
                public const string Secondary = "6464B8";
                public const string Accent = "7878D0";
            }
            public static class Vibrant
            {
                public const string Primary = "5050E0";
                public const string Secondary = "6464F0";
                public const string Accent = "7878FF";
            }
        }
        public static class UI
        {
            public const string Border = "C8C8E8";
            public const string TableHeader = "F0F0FC";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Sand - Warm neutral, real estate
    // ════════════════════════════════════════════════════════════════════════
    public static class Sand
    {
        public static class Text
        {
            public const string Primary = "302820";
            public const string Secondary = "483C30";
            public const string Body = "605040";
            public const string Muted = "887868";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "A89078";
                public const string Secondary = "BCA890";
                public const string Accent = "D0C0A8";
            }
            public static class Vibrant
            {
                public const string Primary = "C89860";
                public const string Secondary = "D8A870";
                public const string Accent = "E8B880";
            }
        }
        public static class UI
        {
            public const string Border = "E0D4C4";
            public const string TableHeader = "F8F4EE";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Ocean - Coastal, wellness
    // ════════════════════════════════════════════════════════════════════════
    public static class Ocean
    {
        public static class Text
        {
            public const string Primary = "0A2030";
            public const string Secondary = "143448";
            public const string Body = "204860";
            public const string Muted = "406888";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "3080A8";
                public const string Secondary = "4098C0";
                public const string Accent = "50B0D8";
            }
            public static class Vibrant
            {
                public const string Primary = "0090D0";
                public const string Secondary = "00A8E8";
                public const string Accent = "00C0FF";
            }
        }
        public static class UI
        {
            public const string Border = "B0D4E8";
            public const string TableHeader = "E8F4FA";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Mint - Financial, trust
    // ════════════════════════════════════════════════════════════════════════
    public static class Mint
    {
        public static class Text
        {
            public const string Primary = "0C2820";
            public const string Secondary = "184038";
            public const string Body = "245850";
            public const string Muted = "408070";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "309878";
                public const string Secondary = "40B090";
                public const string Accent = "50C8A8";
            }
            public static class Vibrant
            {
                public const string Primary = "10B888";
                public const string Secondary = "20D0A0";
                public const string Accent = "30E8B8";
            }
        }
        public static class UI
        {
            public const string Border = "A8E0D0";
            public const string TableHeader = "E8F8F4";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Burgundy - Wine, luxury hospitality
    // ════════════════════════════════════════════════════════════════════════
    public static class Burgundy
    {
        public static class Text
        {
            public const string Primary = "2A0C18";
            public const string Secondary = "441828";
            public const string Body = "5E2438";
            public const string Muted = "884050";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "803848";
                public const string Secondary = "984858";
                public const string Accent = "B05868";
            }
            public static class Vibrant
            {
                public const string Primary = "A02040";
                public const string Secondary = "B83050";
                public const string Accent = "D04060";
            }
        }
        public static class UI
        {
            public const string Border = "E4C0C8";
            public const string TableHeader = "FAF0F2";
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // Graphite - Editorial, publishing
    // ════════════════════════════════════════════════════════════════════════
    public static class Graphite
    {
        public static class Text
        {
            public const string Primary = "1A1A1E";
            public const string Secondary = "2A2A30";
            public const string Body = "3C3C44";
            public const string Muted = "5C5C68";
        }
        public static class Decorative
        {
            public static class Muted
            {
                public const string Primary = "505058";
                public const string Secondary = "686870";
                public const string Accent = "808088";
            }
            public static class Vibrant
            {
                public const string Primary = "404050";
                public const string Secondary = "505060";
                public const string Accent = "606070";
            }
        }
        public static class UI
        {
            public const string Border = "CCCCD0";
            public const string TableHeader = "F2F2F4";
        }
    }
}
