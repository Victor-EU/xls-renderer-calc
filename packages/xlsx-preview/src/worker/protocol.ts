/**
 * The messages the Worker and its host exchange.
 *
 * Kept in its own file so both halves import the same definitions and a change
 * to one is a type error in the other, rather than a message that is silently
 * ignored at runtime.
 */

import type { BindOptions } from '../bind.js';
import type { PreviewSnapshot } from '../snapshot.js';

export interface LoadRequest {
  id: number;
  kind: 'load';
  buf: ArrayBuffer;
  opts: BindOptions;
}

export interface PrecedentsRequest {
  id: number;
  kind: 'precedents';
  sheet: number;
  row: number;
  col: number;
  limit?: number;
}

export type WorkerRequest = LoadRequest | PrecedentsRequest;

export interface LoadResponse {
  id: number;
  kind: 'load';
  snapshot: PreviewSnapshot;
}

export interface PrecedentsResponse {
  id: number;
  kind: 'precedents';
  cells: Array<{ sheet: number; row: number; col: number }>;
}

export interface ErrorResponse {
  id: number;
  kind: 'error';
  message: string;
}

export type WorkerResponse = LoadResponse | PrecedentsResponse | ErrorResponse;
