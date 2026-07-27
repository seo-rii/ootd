use std::mem::size_of;

const CFB_STREAM_SIZE: usize = 4096;
const CFB_FREE_SECTOR: u32 = 0xffff_ffff;
const CFB_END_OF_CHAIN: u32 = 0xffff_fffe;
const CFB_FAT_SECTOR: u32 = 0xffff_fffd;

pub fn compound_file_with_streams(stream_names: &[&str]) -> Vec<u8> {
    compound_file_with_version_and_streams(3, stream_names)
}

#[allow(dead_code)]
pub fn version_4_compound_file_with_streams(stream_names: &[&str]) -> Vec<u8> {
    compound_file_with_version_and_streams(4, stream_names)
}

fn compound_file_with_version_and_streams(major_version: u16, stream_names: &[&str]) -> Vec<u8> {
    assert!(
        stream_names.len() <= 3,
        "fixture supports at most three stream entries"
    );
    let sector_shift = match major_version {
        3 => 9,
        4 => 12,
        _ => panic!("fixture supports CFB major versions 3 and 4"),
    };
    let sector_size = 1_usize << sector_shift;
    let stream_sectors = CFB_STREAM_SIZE / sector_size;
    let fat_sector = 1 + stream_names.len() * stream_sectors;
    let sector_count = fat_sector + 1;
    assert!(
        sector_count <= sector_size / size_of::<u32>(),
        "fixture requires exactly one FAT sector"
    );
    let mut bytes = vec![0_u8; (sector_count + 1) * sector_size];

    bytes[..8].copy_from_slice(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    write_u16(&mut bytes, 24, 0x003e);
    write_u16(&mut bytes, 26, major_version);
    write_u16(&mut bytes, 28, 0xfffe);
    write_u16(&mut bytes, 30, sector_shift);
    write_u16(&mut bytes, 32, 6);
    write_u32(&mut bytes, 40, u32::from(major_version == 4));
    write_u32(&mut bytes, 44, 1);
    write_u32(&mut bytes, 48, 0);
    write_u32(&mut bytes, 56, 4096);
    write_u32(&mut bytes, 60, CFB_END_OF_CHAIN);
    write_u32(&mut bytes, 64, 0);
    write_u32(&mut bytes, 68, CFB_END_OF_CHAIN);
    write_u32(&mut bytes, 72, 0);
    for offset in (76..512).step_by(size_of::<u32>()) {
        write_u32(&mut bytes, offset, CFB_FREE_SECTOR);
    }
    write_u32(&mut bytes, 76, fat_sector as u32);

    let directory_offset = sector_size;
    write_directory_entry(
        &mut bytes,
        directory_offset,
        "Root Entry",
        5,
        CFB_FREE_SECTOR,
        CFB_FREE_SECTOR,
        if stream_names.is_empty() {
            CFB_FREE_SECTOR
        } else {
            1
        },
        CFB_END_OF_CHAIN,
        0,
    );
    for (index, stream_name) in stream_names.iter().enumerate() {
        let start_sector = 1 + index * stream_sectors;
        let entry_offset = directory_offset + (index + 1) * 128;
        write_directory_entry(
            &mut bytes,
            entry_offset,
            stream_name,
            2,
            CFB_FREE_SECTOR,
            if index + 1 < stream_names.len() {
                (index + 2) as u32
            } else {
                CFB_FREE_SECTOR
            },
            CFB_FREE_SECTOR,
            start_sector as u32,
            CFB_STREAM_SIZE as u64,
        );
        let stream_offset = (start_sector + 1) * sector_size;
        if stream_name.eq_ignore_ascii_case("EncryptionInfo") {
            bytes[stream_offset..stream_offset + 8].copy_from_slice(&[4, 0, 4, 0, 0x40, 0, 0, 0]);
        } else if stream_name.eq_ignore_ascii_case("EncryptedPackage") {
            bytes[stream_offset..stream_offset + 8].copy_from_slice(&4096_u64.to_le_bytes());
        }
    }

    let fat_offset = (fat_sector + 1) * sector_size;
    for offset in (fat_offset..fat_offset + sector_size).step_by(size_of::<u32>()) {
        write_u32(&mut bytes, offset, CFB_FREE_SECTOR);
    }
    write_u32(&mut bytes, fat_offset, CFB_END_OF_CHAIN);
    for stream_index in 0..stream_names.len() {
        let start_sector = 1 + stream_index * stream_sectors;
        for sector_offset in 0..stream_sectors {
            let sector = start_sector + sector_offset;
            let next = if sector_offset + 1 == stream_sectors {
                CFB_END_OF_CHAIN
            } else {
                (sector + 1) as u32
            };
            write_u32(&mut bytes, fat_offset + sector * size_of::<u32>(), next);
        }
    }
    write_u32(
        &mut bytes,
        fat_offset + fat_sector * size_of::<u32>(),
        CFB_FAT_SECTOR,
    );
    bytes
}

#[allow(clippy::too_many_arguments)]
fn write_directory_entry(
    bytes: &mut [u8],
    offset: usize,
    name: &str,
    object_type: u8,
    left_sibling: u32,
    right_sibling: u32,
    child: u32,
    start_sector: u32,
    stream_size: u64,
) {
    let name_units = name.encode_utf16().collect::<Vec<_>>();
    assert!(name_units.len() <= 31, "CFB fixture name is too long");
    for (index, unit) in name_units.iter().copied().enumerate() {
        write_u16(bytes, offset + index * size_of::<u16>(), unit);
    }
    write_u16(
        bytes,
        offset + 64,
        ((name_units.len() + 1) * size_of::<u16>()) as u16,
    );
    bytes[offset + 66] = object_type;
    bytes[offset + 67] = 1;
    write_u32(bytes, offset + 68, left_sibling);
    write_u32(bytes, offset + 72, right_sibling);
    write_u32(bytes, offset + 76, child);
    write_u32(bytes, offset + 116, start_sector);
    bytes[offset + 120..offset + 128].copy_from_slice(&stream_size.to_le_bytes());
}

fn write_u16(bytes: &mut [u8], offset: usize, value: u16) {
    bytes[offset..offset + size_of::<u16>()].copy_from_slice(&value.to_le_bytes());
}

fn write_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + size_of::<u32>()].copy_from_slice(&value.to_le_bytes());
}
