# Venue Enrichment Pipeline

Production-grade enrichment pipeline for music venues using:
- Google Places
- SerpAPI
- Local-first processing

## Core Capabilities
- Google Place ID resolution
- Social + metadata enrichment
- Mapotic-ready output formatting
- High-throughput batch processing (10k+ rows)
- Local caching + cost control

## Project Structure
- `src/` → core pipeline scripts
- `config/` → cities, rate limits
- `docs/` → field specs + architecture
- `data/` → runtime inputs/outputs (ignored)

## Quick Start

```bash
python src/enrich_mapotic_places.py \
  --input-xlsx data/input/input.xlsx \
  --output-xlsx data/output/output.xlsx \
  --workers 8 \
  --skip-openai
