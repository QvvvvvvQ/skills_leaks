"""
Palettes - Curated Color Schemes for Professional Documents

Design Principles:
1. Typography.* - Text-safe colors with appropriate saturation for legibility
2. Ornamental.Subtle.* - Muted decorative tones for professional contexts
3. Ornamental.Bold.* - Saturated accents for emphasis and visual impact
4. Interface.* - Structural elements (borders, backgrounds)

NEVER apply Ornamental.Bold.* to body text - readability degradation guaranteed.

Palette Selection Guide:
  Verdant - Deep forest greens, editorial aesthetic
  Graphite - Monochromatic zen, contemplative
  Executive - Professional blue, corporate identity
  Terracotta - Warm brown/olive, organic feel
  Nordic - Cool gray-blue, minimalist tech
  Blush - Dusty rose/cream, refined elegance
  Academic - Navy/burgundy, scholarly gravitas
  Maritime - Ocean blue/sand, wellness themes
  Canopy - Olive/moss, sustainability focus
  Industrial - Charcoal/rust, urban strength
  Dune - Ochre/gold, warm regional character
  Regal - Formal red/gold, authority themes
  Luxe - Black/gold, luxury positioning
  Signal - Yellow/black, high-energy promotional
  Nova - Purple/indigo, innovation/tech
  Clinical - Teal/white, medical/clean
  Minimal - Pure monochrome, stark editorial
  Citrus - Orange/green, vitality/wellness
  Neon - Slate/electric, cyber/analytical
  Coral - Warm pink, lifestyle/fashion
  Mint - Green/slate, financial trust
"""

# ============================================================================
# Verdant - Deep forest greens, editorial aesthetic (replaces Sage)
# ============================================================================
Verdant = {
    'Typography': {
        'Heading': '1F3D2B',      # Deep forest
        'Subheading': '2E5A3D',   # Woodland
        'Paragraph': '436B4F',    # Fern
        'Caption': '7A9E85',      # Moss light
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '4D7A5C',
            'Secondary': '6B9A8A',
            'Highlight': '89B89A',
        },
        'Bold': {
            'Primary': '2D6A4F',
            'Secondary': '40916C',
            'Highlight': '52B788',
        },
    },
    'Interface': {
        'Divider': 'B7E4C7',
        'Surface': 'D8F3DC',
    },
}

# ============================================================================
# Graphite - Monochromatic zen, contemplative
# ============================================================================
Graphite = {
    'Typography': {
        'Heading': '18181B',
        'Subheading': '27272A',
        'Paragraph': '3F3F46',
        'Caption': '71717A',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '52525B',
            'Secondary': '6B6B74',
            'Highlight': '8B8B96',
        },
        'Bold': {
            'Primary': '3F3F46',
            'Secondary': '52525B',
            'Highlight': 'A1A1AA',
        },
    },
    'Interface': {
        'Divider': 'D4D4D8',
        'Surface': 'F4F4F5',
    },
}

# ============================================================================
# Executive - Professional blue, corporate identity
# ============================================================================
Executive = {
    'Typography': {
        'Heading': '0F172A',
        'Subheading': '1E3A5F',
        'Paragraph': '334155',
        'Caption': '64748B',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '475569',
            'Secondary': '64748B',
            'Highlight': '94A3B8',
        },
        'Bold': {
            'Primary': '1D4ED8',
            'Secondary': '3B82F6',
            'Highlight': '60A5FA',
        },
    },
    'Interface': {
        'Divider': 'CBD5E1',
        'Surface': 'F1F5F9',
    },
}

# ============================================================================
# Terracotta - Warm brown/olive, organic feel
# ============================================================================
Terracotta = {
    'Typography': {
        'Heading': '422006',
        'Subheading': '78350F',
        'Paragraph': '92400E',
        'Caption': 'B45309',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '78350F',
            'Secondary': '92400E',
            'Highlight': 'D97706',
        },
        'Bold': {
            'Primary': 'B45309',
            'Secondary': 'D97706',
            'Highlight': 'F59E0B',
        },
    },
    'Interface': {
        'Divider': 'FED7AA',
        'Surface': 'FFF7ED',
    },
}

# ============================================================================
# Nordic - Cool gray-blue, minimalist tech (replaces Fjord)
# ============================================================================
Nordic = {
    'Typography': {
        'Heading': '1E293B',
        'Subheading': '334155',
        'Paragraph': '475569',
        'Caption': '94A3B8',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '64748B',
            'Secondary': '94A3B8',
            'Highlight': 'CBD5E1',
        },
        'Bold': {
            'Primary': '0284C7',
            'Secondary': '0EA5E9',
            'Highlight': '38BDF8',
        },
    },
    'Interface': {
        'Divider': 'E2E8F0',
        'Surface': 'F8FAFC',
    },
}

# ============================================================================
# Blush - Dusty rose/cream, refined elegance (replaces Provence)
# ============================================================================
Blush = {
    'Typography': {
        'Heading': '4A2535',
        'Subheading': '6B3A4D',
        'Paragraph': '8B5A6B',
        'Caption': 'A87A8B',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': 'C4A0AB',
            'Secondary': 'D4B5BF',
            'Highlight': 'E8D0D8',
        },
        'Bold': {
            'Primary': 'BE185D',
            'Secondary': 'DB2777',
            'Highlight': 'EC4899',
        },
    },
    'Interface': {
        'Divider': 'FECDD3',
        'Surface': 'FFF1F2',
    },
}

# ============================================================================
# Academic - Navy/burgundy, scholarly gravitas (replaces Scholar)
# ============================================================================
Academic = {
    'Typography': {
        'Heading': '1E1B4B',
        'Subheading': '312E81',
        'Paragraph': '3730A3',
        'Caption': '6366F1',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '4338CA',
            'Secondary': '6366F1',
            'Highlight': '818CF8',
        },
        'Bold': {
            'Primary': '7C2D12',
            'Secondary': '9A3412',
            'Highlight': 'C2410C',
        },
    },
    'Interface': {
        'Divider': 'C7D2FE',
        'Surface': 'EEF2FF',
    },
}

# ============================================================================
# Maritime - Ocean blue/sand, wellness themes (replaces Aegean)
# ============================================================================
Maritime = {
    'Typography': {
        'Heading': '0C4A6E',
        'Subheading': '075985',
        'Paragraph': '0369A1',
        'Caption': '0284C7',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '38BDF8',
            'Secondary': '7DD3FC',
            'Highlight': 'BAE6FD',
        },
        'Bold': {
            'Primary': '0284C7',
            'Secondary': '0EA5E9',
            'Highlight': '38BDF8',
        },
    },
    'Interface': {
        'Divider': 'E0F2FE',
        'Surface': 'F0F9FF',
    },
}

# ============================================================================
# Canopy - Olive/moss, sustainability focus (replaces Evergreen)
# ============================================================================
Canopy = {
    'Typography': {
        'Heading': '14532D',
        'Subheading': '166534',
        'Paragraph': '15803D',
        'Caption': '16A34A',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '22C55E',
            'Secondary': '4ADE80',
            'Highlight': '86EFAC',
        },
        'Bold': {
            'Primary': '15803D',
            'Secondary': '16A34A',
            'Highlight': '22C55E',
        },
    },
    'Interface': {
        'Divider': 'BBF7D0',
        'Surface': 'DCFCE7',
    },
}

# ============================================================================
# Industrial - Charcoal/rust, urban strength (replaces Foundry)
# ============================================================================
Industrial = {
    'Typography': {
        'Heading': '1C1917',
        'Subheading': '292524',
        'Paragraph': '44403C',
        'Caption': '78716C',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '57534E',
            'Secondary': '78716C',
            'Highlight': 'A8A29E',
        },
        'Bold': {
            'Primary': 'B91C1C',
            'Secondary': 'DC2626',
            'Highlight': 'EF4444',
        },
    },
    'Interface': {
        'Divider': 'D6D3D1',
        'Surface': 'F5F5F4',
    },
}

# ============================================================================
# Dune - Ochre/gold, warm regional character (replaces Sahara)
# ============================================================================
Dune = {
    'Typography': {
        'Heading': '451A03',
        'Subheading': '78350F',
        'Paragraph': '92400E',
        'Caption': 'B45309',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': 'D97706',
            'Secondary': 'F59E0B',
            'Highlight': 'FBBF24',
        },
        'Bold': {
            'Primary': 'B45309',
            'Secondary': 'D97706',
            'Highlight': 'F59E0B',
        },
    },
    'Interface': {
        'Divider': 'FDE68A',
        'Surface': 'FFFBEB',
    },
}

# ============================================================================
# Regal - Formal red/gold, authority themes (replaces Cardinal)
# ============================================================================
Regal = {
    'Typography': {
        'Heading': '450A0A',
        'Subheading': '7F1D1D',
        'Paragraph': '991B1B',
        'Caption': 'B91C1C',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': 'DC2626',
            'Secondary': 'EF4444',
            'Highlight': 'F87171',
        },
        'Bold': {
            'Primary': 'CA8A04',
            'Secondary': 'EAB308',
            'Highlight': 'FACC15',
        },
    },
    'Interface': {
        'Divider': 'FECACA',
        'Surface': 'FEF2F2',
    },
}

# ============================================================================
# Luxe - Black/gold, luxury positioning (replaces Opulent)
# ============================================================================
Luxe = {
    'Typography': {
        'Heading': '0A0A0A',
        'Subheading': '171717',
        'Paragraph': '262626',
        'Caption': '525252',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '404040',
            'Secondary': '737373',
            'Highlight': 'A3A3A3',
        },
        'Bold': {
            'Primary': 'A16207',
            'Secondary': 'CA8A04',
            'Highlight': 'EAB308',
        },
    },
    'Interface': {
        'Divider': 'D4D4D4',
        'Surface': 'FAFAFA',
    },
}

# ============================================================================
# Signal - Yellow/black, high-energy promotional (replaces Vivid)
# ============================================================================
Signal = {
    'Typography': {
        'Heading': '0A0A0A',
        'Subheading': '171717',
        'Paragraph': '262626',
        'Caption': '525252',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': 'A16207',
            'Secondary': 'CA8A04',
            'Highlight': 'EAB308',
        },
        'Bold': {
            'Primary': 'EAB308',
            'Secondary': 'FACC15',
            'Highlight': 'FDE047',
        },
    },
    'Interface': {
        'Divider': 'FEF08A',
        'Surface': 'FEFCE8',
    },
}

# ============================================================================
# Nova - Purple/indigo, innovation/tech (replaces Catalyst)
# ============================================================================
Nova = {
    'Typography': {
        'Heading': '2E1065',
        'Subheading': '4C1D95',
        'Paragraph': '5B21B6',
        'Caption': '7C3AED',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '8B5CF6',
            'Secondary': 'A78BFA',
            'Highlight': 'C4B5FD',
        },
        'Bold': {
            'Primary': '7C3AED',
            'Secondary': '8B5CF6',
            'Highlight': 'A78BFA',
        },
    },
    'Interface': {
        'Divider': 'DDD6FE',
        'Surface': 'F5F3FF',
    },
}

# ============================================================================
# Clinical - Teal/white, medical/clean (replaces Pristine)
# ============================================================================
Clinical = {
    'Typography': {
        'Heading': '134E4A',
        'Subheading': '115E59',
        'Paragraph': '0F766E',
        'Caption': '14B8A6',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '2DD4BF',
            'Secondary': '5EEAD4',
            'Highlight': '99F6E4',
        },
        'Bold': {
            'Primary': '0D9488',
            'Secondary': '14B8A6',
            'Highlight': '2DD4BF',
        },
    },
    'Interface': {
        'Divider': 'CCFBF1',
        'Surface': 'F0FDFA',
    },
}

# ============================================================================
# Minimal - Pure monochrome, stark editorial (replaces Zen)
# ============================================================================
Minimal = {
    'Typography': {
        'Heading': '0A0A0A',
        'Subheading': '171717',
        'Paragraph': '262626',
        'Caption': '737373',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '404040',
            'Secondary': '525252',
            'Highlight': '737373',
        },
        'Bold': {
            'Primary': '171717',
            'Secondary': '262626',
            'Highlight': '404040',
        },
    },
    'Interface': {
        'Divider': 'E5E5E5',
        'Surface': 'FAFAFA',
    },
}

# ============================================================================
# Citrus - Orange/green, vitality/wellness (replaces Radiant)
# ============================================================================
Citrus = {
    'Typography': {
        'Heading': '7C2D12',
        'Subheading': '9A3412',
        'Paragraph': 'C2410C',
        'Caption': 'EA580C',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': 'F97316',
            'Secondary': 'FB923C',
            'Highlight': 'FDBA74',
        },
        'Bold': {
            'Primary': '16A34A',
            'Secondary': '22C55E',
            'Highlight': '4ADE80',
        },
    },
    'Interface': {
        'Divider': 'FED7AA',
        'Surface': 'FFF7ED',
    },
}

# ============================================================================
# Neon - Slate/electric, cyber/analytical (replaces Eclipse)
# ============================================================================
Neon = {
    'Typography': {
        'Heading': '0F172A',
        'Subheading': '1E293B',
        'Paragraph': '334155',
        'Caption': '64748B',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '475569',
            'Secondary': '64748B',
            'Highlight': '94A3B8',
        },
        'Bold': {
            'Primary': '06B6D4',
            'Secondary': '22D3EE',
            'Highlight': 'F472B6',
        },
    },
    'Interface': {
        'Divider': 'CBD5E1',
        'Surface': 'F8FAFC',
    },
}

# ============================================================================
# Coral - Warm pink, lifestyle/fashion (replaces Lagoon)
# ============================================================================
Coral = {
    'Typography': {
        'Heading': '831843',
        'Subheading': '9D174D',
        'Paragraph': 'BE185D',
        'Caption': 'DB2777',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': 'EC4899',
            'Secondary': 'F472B6',
            'Highlight': 'F9A8D4',
        },
        'Bold': {
            'Primary': 'F43F5E',
            'Secondary': 'FB7185',
            'Highlight': 'FDA4AF',
        },
    },
    'Interface': {
        'Divider': 'FBCFE8',
        'Surface': 'FDF2F8',
    },
}

# ============================================================================
# Mint - Green/slate, financial trust (replaces Sterling)
# ============================================================================
Mint = {
    'Typography': {
        'Heading': '064E3B',
        'Subheading': '065F46',
        'Paragraph': '047857',
        'Caption': '059669',
    },
    'Ornamental': {
        'Subtle': {
            'Primary': '10B981',
            'Secondary': '34D399',
            'Highlight': '6EE7B7',
        },
        'Bold': {
            'Primary': '047857',
            'Secondary': '059669',
            'Highlight': '10B981',
        },
    },
    'Interface': {
        'Divider': 'A7F3D0',
        'Surface': 'ECFDF5',
    },
}

# ============================================================================
# Registry - All palettes indexed
# ============================================================================
REGISTRY = {
    'Verdant': Verdant,
    'Graphite': Graphite,
    'Executive': Executive,
    'Terracotta': Terracotta,
    'Nordic': Nordic,
    'Blush': Blush,
    'Academic': Academic,
    'Maritime': Maritime,
    'Canopy': Canopy,
    'Industrial': Industrial,
    'Dune': Dune,
    'Regal': Regal,
    'Luxe': Luxe,
    'Signal': Signal,
    'Nova': Nova,
    'Clinical': Clinical,
    'Minimal': Minimal,
    'Citrus': Citrus,
    'Neon': Neon,
    'Coral': Coral,
    'Mint': Mint,
}


def to_css_format(hex_value: str) -> str:
    """Convert bare hex to CSS notation."""
    return f'#{hex_value}'


def fetch_palette(name: str) -> dict:
    """Retrieve palette by name. Raises KeyError if not found."""
    return REGISTRY[name]


def fetch_typography(palette_name: str, level: str = 'Paragraph') -> str:
    """Get typography color. level: Heading, Subheading, Paragraph, Caption"""
    return REGISTRY[palette_name]['Typography'][level]


def fetch_ornamental(palette_name: str, intensity: str = 'Subtle', level: str = 'Primary') -> str:
    """Get ornamental color. intensity: Subtle/Bold. level: Primary/Secondary/Highlight"""
    return REGISTRY[palette_name]['Ornamental'][intensity][level]
