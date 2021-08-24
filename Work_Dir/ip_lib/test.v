// 
// ******************************************************************************
// *                                                                            *
// *                   Copyright (C) 2004-2015, Nangate Inc.                    *
// *                           All rights reserved.                             *
// *                                                                            *
// * Nangate and the Nangate logo are trademarks of Nangate Inc.                *
// *                                                                            *
// * All trademarks, logos, software marks, and trade names (collectively the   *
// * "Marks") in this program are proprietary to Nangate or other respective    *
// * owners that have granted Nangate the right and license to use such Marks.  *
// * You are not permitted to use the Marks without the prior written consent   *
// * of Nangate or such third party that may own the Marks.                     *
// *                                                                            *
// * This file has been provided pursuant to a License Agreement containing     *
// * restrictions on its use. This file contains valuable trade secrets and     *
// * proprietary information of Nangate Inc., and is protected by U.S. and      *
// * international laws and/or treaties.                                        *
// *                                                                            *
// * The copyright notice(s) in this file does not indicate actual or intended  *
// * publication of this file.                                                  *
// *                                                                            *
// *    NGLibraryCharacterizer, Development_version_64 - build 201511030503     *
// *                                                                            *
// ******************************************************************************
// 
// 
// Running on us19.nangate.us for user Bernardo Predebon Toffoli Culau (btc).
// Local time is now Tue, 17 Nov 2015, 15:00:35.
// Main process id is 20374.
// 
// * Default delays
//   * comb. path delay        : 0.1
//   * seq. path delay         : 0.1
//   * delay cells             : 0.1
//   * timing checks           : 0.1
// 
// * NTC Setup
//   * Export NTC sections     : true
//   * Combine setup / hold    : true
//   * Combine recovery/removal: true
// 
// * Extras
//   * Export `celldefine      : false
//   * Export `timescale       : -
// 

module LS_HLEN_X1_RVT_30 (A, ISOLN, Z);
  input A;
  input ISOLN;
  output Z;

  and(Z, A, ISOLN);
endmodule

module LS_HLEN_X2_RVT_30 (A, ISOLN, Z);
  input A;
  input ISOLN;
  output Z;

  and(Z, A, ISOLN);
endmodule

module LS_HLEN_X4_RVT_30 (A, ISOLN, Z);
  input A;
  input ISOLN;
  output Z;

  and(Z, A, ISOLN);
endmodule

module LS_HLEN_X8_RVT_30 (A, ISOLN, Z);
  input A;
  input ISOLN;
  output Z;

  and(Z, A, ISOLN);
endmodule

module LS_HL_X1_RVT_30 (A, Z);
  input A;
  output Z;

  buf(Z, A);
endmodule

module LS_HL_X2_RVT_30 (A, Z);
  input A;
  output Z;

  buf(Z, A);
endmodule

module LS_HL_X4_RVT_30 (A, Z);
  input A;
  output Z;

  buf(Z, A);
endmodule

module LS_HL_X8_RVT_30 (A, Z);
  input A;
  output Z;

  buf(Z, A);
endmodule

module LS_LHEN_X1_RVT_30 (A, ISOLN, Z);
  input A;
  input ISOLN;
  output Z;

  or(Z, A, i_0);
  not(i_0, ISOLN);
endmodule

module LS_LHEN_X2_RVT_30 (A, ISOLN, Z);
  input A;
  input ISOLN;
  output Z;

  or(Z, A, i_0);
  not(i_0, ISOLN);
endmodule

module LS_LHEN_X4_RVT_30 (A, ISOLN, Z);
  input A;
  input ISOLN;
  output Z;

  or(Z, A, i_0);
  not(i_0, ISOLN);
endmodule

module LS_LHEN_X8_RVT_30 (A, ISOLN, Z);
  input A;
  input ISOLN;
  output Z;

  or(Z, A, i_0);
  not(i_0, ISOLN);
endmodule

module LS_LH_X1_RVT_30 (A, Z);
  input A;
  output Z;

  buf(Z, A);
endmodule

module LS_LH_X2_RVT_30 (A, Z);
  input A;
  output Z;

  buf(Z, A);
endmodule

module LS_LH_X4_RVT_30 (A, Z);
  input A;
  output Z;

  buf(Z, A);
endmodule

module LS_LH_X8_RVT_30 (A, Z);
  input A;
  output Z;

  buf(Z, A);
endmodule

`ifdef TETRAMAX
`else
  primitive ng_xbuf (o, i, d);
	output o;
	input i, d;
	table
	// i   d   : o
	   0   1   : 0 ;
	   1   1   : 1 ;
	   x   1   : 1 ;
	endtable
  endprimitive
`endif
//
// End of file
//